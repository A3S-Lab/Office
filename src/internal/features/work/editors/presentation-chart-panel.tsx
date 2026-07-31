import { Plus, Trash2, X } from 'lucide-react';
import { Button, IconButton } from '../../../design-system/primitives';
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
import { handleOfficeTaskPaneKeyDown } from './office-task-pane';

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
}: {
  chart: WorkSlideChart;
  onChange: (chart: WorkSlideChart) => void;
  onDelete: () => void;
  onClose: () => void;
}) {
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
      className="work-presentation-chart-panel"
      aria-label="演示图表数据"
      onKeyDown={(event) => handleOfficeTaskPaneKeyDown(event, onClose)}
    >
      <header>
        <div>
          <strong>图表数据</strong>
          <span>配置数据、布局与系列</span>
        </div>
        <div>
          <Button tone="danger" aria-label="删除演示图表" onClick={onDelete}>
            <Trash2 size={13} />
            删除图表
          </Button>
          <IconButton
            className="close"
            label="关闭演示图表数据"
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
          <span>类型</span>
          <OfficeSelect
            ariaLabel="演示图表类型"
            value={chart.type}
            options={CHART_TYPES.map((type) => ({
              value: type,
              label: presentationChartTypeLabel(type),
            }))}
            onValueChange={(type) =>
              onChange(
                withPresentationChartType(chart, type as WorkSlideChartType),
              )
            }
          />
        </div>
        <div className="work-office-field">
          <span>标题</span>
          <CommittedOfficeTextField
            aria-label="演示图表标题"
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
            <span>孔径</span>
            <CommittedOfficeNumberField
              ariaLabel="圆环孔径"
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
            <span>样式</span>
            <OfficeSelect
              ariaLabel="雷达图样式"
              value={chart.radarStyle ?? 'standard'}
              options={[
                { value: 'standard', label: '标准' },
                { value: 'marker', label: '带数据标记' },
                { value: 'filled', label: '填充' },
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
            <span>散点样式</span>
            <OfficeSelect
              ariaLabel="演示散点图样式"
              value={normalizePresentationScatterStyle(chart.scatterStyle)}
              options={[
                { value: 'marker', label: '仅数据标记' },
                { value: 'line', label: '直线' },
                { value: 'lineMarker', label: '直线和数据标记' },
                { value: 'smooth', label: '平滑线' },
                { value: 'smoothMarker', label: '平滑线和数据标记' },
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
              <span>气泡缩放</span>
              <CommittedOfficeNumberField
                ariaLabel="演示气泡图缩放"
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
              <span>大小表示</span>
              <OfficeSelect
                ariaLabel="演示气泡大小表示"
                value={normalizePresentationBubbleSizeRepresents(
                  chart.bubbleSizeRepresents,
                )}
                options={[
                  { value: 'area', label: '面积' },
                  { value: 'width', label: '宽度' },
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
              <span>负气泡</span>
              <OfficeCheckbox
                className="work-presentation-chart-check-control"
                ariaLabel="显示负气泡"
                checked={chart.showNegativeBubbles === true}
                onCheckedChange={(showNegativeBubbles) =>
                  onChange({ ...chart, showNegativeBubbles })
                }
              >
                显示
              </OfficeCheckbox>
            </div>
          </>
        )}
        <div className="work-office-field categories">
          <span>{numericXAxis ? 'X 值' : '分类'}（每行一项）</span>
          <CommittedOfficeTextArea
            aria-label={numericXAxis ? '演示图表 X 值' : '演示图表分类'}
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
                <legend>系列 {index + 1}</legend>
                <CommittedOfficeTextField
                  aria-label={`演示图表系列 ${index + 1} 名称`}
                  value={series.name}
                  formatValue={(name) => name}
                  parseValue={(draft) => draft.trim().slice(0, 255)}
                  onValueCommit={(name) => updateSeries(index, { name })}
                />
                <CommittedOfficeTextArea
                  aria-label={`演示图表系列 ${index + 1} ${numericXAxis ? 'Y 值' : '数据'}`}
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
                    aria-label={`演示气泡图系列 ${index + 1} 大小`}
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
                  label={`删除演示图表系列 ${index + 1}`}
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
            aria-label="添加图表系列"
            onClick={() =>
              onChange({
                ...chart,
                series: [...chart.series, createPresentationChartSeries(chart)],
              })
            }
          >
            <Plus size={13} />
            添加系列
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
