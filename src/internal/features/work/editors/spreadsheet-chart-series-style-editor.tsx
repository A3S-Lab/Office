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
      aria-label={`系列 ${seriesNumber} 外观设置`}
    >
      <OfficeCheckbox
        className="check enable-series-style"
        ariaLabel={`系列 ${seriesNumber} 使用自定义外观`}
        checked={Boolean(value)}
        onCheckedChange={(checked) => onChange(checked ? defaults : undefined)}
      >
        自定义系列外观
      </OfficeCheckbox>
      {value && (
        <div>
          <div className="work-office-field">
            <span>填充颜色</span>
            <OfficeColorPicker
              ariaLabel={`系列 ${seriesNumber} 填充颜色`}
              value={style.fillColor ?? '#4F6BED'}
              onValueChange={(fillColor) => update({ fillColor })}
            />
          </div>
          <div className="work-office-field">
            <span>填充透明度（%）</span>
            <CommittedOfficeNumberField
              ariaLabel={`系列 ${seriesNumber} 填充透明度`}
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
            <span>线条颜色</span>
            <OfficeColorPicker
              ariaLabel={`系列 ${seriesNumber} 线条颜色`}
              value={style.lineColor ?? '#4F6BED'}
              onValueChange={(lineColor) => update({ lineColor })}
            />
          </div>
          <div className="work-office-field">
            <span>线条宽度（磅）</span>
            <CommittedOfficeNumberField
              ariaLabel={`系列 ${seriesNumber} 线条宽度`}
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
            <span>线条虚线</span>
            <OfficeSelect
              ariaLabel={`系列 ${seriesNumber} 线条虚线`}
              value={style.lineDash ?? 'solid'}
              options={[
                { value: 'solid', label: '实线' },
                { value: 'dash', label: '虚线' },
                { value: 'dot', label: '点线' },
                { value: 'dashDot', label: '点划线' },
              ]}
              onValueChange={(lineDash) =>
                update({ lineDash: lineDash as WorkSpreadsheetChartLineDash })
              }
            />
          </div>
          {supportsMarkers && (
            <>
              <div className="work-office-field">
                <span>数据标记符号</span>
                <OfficeSelect
                  ariaLabel={`系列 ${seriesNumber} 数据标记符号`}
                  value={marker.symbol ?? 'circle'}
                  options={[
                    { value: 'none', label: '无' },
                    { value: 'circle', label: '圆形' },
                    { value: 'square', label: '方形' },
                    { value: 'diamond', label: '菱形' },
                    { value: 'triangle', label: '三角形' },
                    { value: 'plus', label: '加号' },
                    { value: 'x', label: '叉号' },
                    { value: 'star', label: '星形' },
                  ]}
                  onValueChange={(symbol) =>
                    updateMarker({
                      symbol: symbol as WorkSpreadsheetChartMarkerSymbol,
                    })
                  }
                />
              </div>
              <div className="work-office-field">
                <span>数据标记大小（磅）</span>
                <CommittedOfficeNumberField
                  ariaLabel={`系列 ${seriesNumber} 数据标记大小`}
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
                <span>数据标记填充</span>
                <OfficeColorPicker
                  ariaLabel={`系列 ${seriesNumber} 数据标记填充颜色`}
                  value={marker.fillColor ?? '#FFFFFF'}
                  onValueChange={(fillColor) => updateMarker({ fillColor })}
                />
              </div>
              <div className="work-office-field">
                <span>数据标记轮廓</span>
                <OfficeColorPicker
                  ariaLabel={`系列 ${seriesNumber} 数据标记轮廓颜色`}
                  value={marker.lineColor ?? style.lineColor ?? '#4F6BED'}
                  onValueChange={(lineColor) => updateMarker({ lineColor })}
                />
              </div>
            </>
          )}
        </div>
      )}
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
