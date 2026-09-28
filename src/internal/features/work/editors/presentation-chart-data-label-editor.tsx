import { officeMessage } from '../../../i18n/office-locale';
import {
  normalizePresentationChartDataLabelPosition,
  presentationChartDataLabelPositionLabel,
  presentationChartDataLabelPositions,
} from '../work-presentation-charts';
import type {
  WorkSlideChartDataLabelPosition,
  WorkSlideChartDataLabels,
  WorkSlideChartType,
} from '../work-types';
import {
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export function PresentationChartDataLabelEditor({
  chartType,
  value,
  onChange,
}: {
  chartType: WorkSlideChartType;
  value?: WorkSlideChartDataLabels;
  onChange: (value: WorkSlideChartDataLabels | undefined) => void;
}) {
  const messages = useOfficeMessages();
  const change = (patch: Partial<WorkSlideChartDataLabels>) =>
    onChange({ ...value, ...patch });
  return (
    <section
      className="work-presentation-chart-data-labels"
      aria-label={officeMessage(messages, 'presentation.chart.dataLabel.settingsAria')}
    >
      <OfficeCheckbox
        className="check enable-data-labels"
        ariaLabel={officeMessage(messages, 'presentation.chart.dataLabel.showAria')}
        checked={value !== undefined}
        onCheckedChange={(checked) =>
          onChange(checked ? { showValue: true } : undefined)
        }
      >
        {officeMessage(messages, 'presentation.chart.dataLabel.show')}
      </OfficeCheckbox>
      {value && (
        <div>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(messages, 'presentation.chart.dataLabel.showValueAria')}
            checked={value.showValue === true}
            onCheckedChange={(showValue) => change({ showValue })}
          >
            {officeMessage(messages, 'presentation.chart.dataLabel.showValue')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(messages, 'presentation.chart.dataLabel.showCategoryAria')}
            checked={value.showCategoryName === true}
            onCheckedChange={(showCategoryName) => change({ showCategoryName })}
          >
            {officeMessage(messages, 'presentation.chart.dataLabel.showCategory')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(messages, 'presentation.chart.dataLabel.showSeriesAria')}
            checked={value.showSeriesName === true}
            onCheckedChange={(showSeriesName) => change({ showSeriesName })}
          >
            {officeMessage(messages, 'presentation.chart.dataLabel.showSeries')}
          </OfficeCheckbox>
          {(chartType === 'pie' || chartType === 'doughnut') && (
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(messages, 'presentation.chart.dataLabel.showPercentAria')}
              checked={value.showPercentage === true}
              onCheckedChange={(showPercentage) => change({ showPercentage })}
            >
              {officeMessage(messages, 'presentation.chart.dataLabel.showPercent')}
            </OfficeCheckbox>
          )}
          {chartType === 'bubble' && (
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(messages, 'presentation.chart.dataLabel.showBubbleAria')}
              checked={value.showBubbleSize === true}
              onCheckedChange={(showBubbleSize) => change({ showBubbleSize })}
            >
              {officeMessage(messages, 'presentation.chart.dataLabel.showBubble')}
            </OfficeCheckbox>
          )}
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.chart.dataLabel.position')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.chart.dataLabel.positionAria')}
              value={normalizePresentationChartDataLabelPosition(
                value.position,
                chartType,
              )}
              options={presentationChartDataLabelPositions(chartType).map(
                (position) => ({
                  value: position,
                  label: presentationChartDataLabelPositionLabel(position, messages),
                }),
              )}
              onValueChange={(position) =>
                change({
                  position: position as WorkSlideChartDataLabelPosition,
                })
              }
            />
          </div>
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.chart.dataLabel.separator')}</span>
            <OfficeTextField
              aria-label={officeMessage(messages, 'presentation.chart.dataLabel.separatorAria')}
              value={value.separator ?? ', '}
              maxLength={64}
              onChange={(event) => change({ separator: event.target.value })}
            />
          </div>
        </div>
      )}
    </section>
  );
}
