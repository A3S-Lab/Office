import { officeMessage } from '../../../i18n/office-locale';
import { useOfficeMessages } from './office-messages-context';
import { withPresentationChartSeriesAnalysis } from '../work-presentation-charts';
import { workSpreadsheetChartSupportsSeriesAnalysis } from '../work-spreadsheet-chart-layout';
import {
  type WorkSlideChart,
  workSpreadsheetChartSupportsErrorBars,
  workSpreadsheetChartSupportsTrendlines,
} from '../work-types';
import { SpreadsheetErrorBarEditor } from './spreadsheet-error-bar-editor';
import { SpreadsheetTrendlineEditor } from './spreadsheet-trendline-editor';

export function PresentationChartSeriesAnalysisEditor({
  chart,
  seriesIndex,
  onChange,
}: {
  chart: WorkSlideChart;
  seriesIndex: number;
  onChange: (chart: WorkSlideChart) => void;
}) {
  const series = chart.series[seriesIndex];
  if (!series) return null;
  if (!workSpreadsheetChartSupportsSeriesAnalysis(chart)) return null;
  const supportsErrorBars = workSpreadsheetChartSupportsErrorBars(chart.type);
  const supportsTrendlines = workSpreadsheetChartSupportsTrendlines(chart.type);
  if (!supportsErrorBars && !supportsTrendlines) return null;
  const messages = useOfficeMessages();
  const seriesNumber = seriesIndex + 1;

  return (
    <section
      className="work-presentation-chart-series-analysis"
      aria-label={officeMessage(
        messages,
        'presentation.chart.seriesAnalysisAria',
        { number: String(seriesNumber) },
      )}
    >
      <header>
        <strong>
          {officeMessage(messages, 'presentation.chart.seriesAnalysisTitle')}
        </strong>
        <span>
          {officeMessage(messages, 'presentation.chart.seriesAnalysisHint')}
        </span>
      </header>
      {supportsTrendlines && (
        <SpreadsheetTrendlineEditor
          seriesNumber={seriesNumber}
          trendlines={series.trendlines ?? []}
          onChange={(trendlines) =>
            onChange(
              withPresentationChartSeriesAnalysis(chart, seriesIndex, {
                trendlines,
                errorBars: series.errorBars,
              }),
            )
          }
        />
      )}
      {supportsErrorBars && (
        <SpreadsheetErrorBarEditor
          chartType={chart.type}
          seriesNumber={seriesNumber}
          errorBars={series.errorBars ?? []}
          customInput="values"
          onChange={(errorBars) =>
            onChange(
              withPresentationChartSeriesAnalysis(chart, seriesIndex, {
                errorBars,
                trendlines: series.trendlines,
              }),
            )
          }
        />
      )}
    </section>
  );
}
