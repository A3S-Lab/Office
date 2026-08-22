import { Plus, Trash2 } from 'lucide-react';
import type { Dispatch, ReactNode, SetStateAction } from 'react';
import {
  Button,
  IconButton,
  SegmentedControl,
} from '../../../design-system/primitives';
import {
  MAX_XLSX_GRADIENT_STOPS,
  type XlsxGradientFill,
  type XlsxGradientStop,
} from '../work-xlsx-gradient-fill';
import {
  xlsxPatternFillTypes,
  type XlsxPatternFill,
  type XlsxPatternFillType,
} from '../work-xlsx-pattern-fill';
import type { SpreadsheetCellFormatPatch } from './spreadsheet-cell-format';
import {
  spreadsheetFormatCellsActiveFill,
  type SpreadsheetFormatCellsDialogSource,
  type SpreadsheetFormatCellsDraft,
  type SpreadsheetFormatCellsDraftErrors,
  type SpreadsheetFormatCellsFillDraft,
  type SpreadsheetFormatCellsTouched,
} from './spreadsheet-format-cells-dialog-model';
import { SpreadsheetFormatCellsFillPreview } from './spreadsheet-format-cells-fill-preview';
import {
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';

interface FillPanelProps {
  source: SpreadsheetFormatCellsDialogSource;
  draft: SpreadsheetFormatCellsDraft;
  errors: SpreadsheetFormatCellsDraftErrors;
  touched: SpreadsheetFormatCellsTouched;
  setDraft: Dispatch<SetStateAction<SpreadsheetFormatCellsDraft>>;
  touch: (field: keyof SpreadsheetCellFormatPatch) => void;
}

const fillModeItems = [
  { id: 'none', label: '无填充' },
  { id: 'solid', label: '纯色' },
  { id: 'pattern', label: '图案' },
  { id: 'gradient', label: '渐变' },
] as const;

const gradientTypeOptions = [
  { value: 'linear', label: '线性' },
  { value: 'path', label: '路径' },
] as const satisfies readonly OfficeSelectOption<XlsxGradientFill['type']>[];

const patternLabels: Record<XlsxPatternFillType, string> = {
  darkDown: '深色下斜线',
  darkGray: '深灰',
  darkGrid: '深色网格',
  darkHorizontal: '深色横线',
  darkTrellis: '深色菱形网格',
  darkUp: '深色上斜线',
  darkVertical: '深色竖线',
  gray0625: '6.25% 灰度',
  gray125: '12.5% 灰度',
  lightDown: '浅色下斜线',
  lightGray: '浅灰',
  lightGrid: '浅色网格',
  lightHorizontal: '浅色横线',
  lightTrellis: '浅色菱形网格',
  lightUp: '浅色上斜线',
  lightVertical: '浅色竖线',
  mediumGray: '中灰',
};

const patternOptions = xlsxPatternFillTypes.map((value) => ({
  label: patternLabels[value],
  value,
}));

export function SpreadsheetFormatCellsFillPanel({
  source,
  draft,
  errors,
  touched,
  setDraft,
  touch,
}: FillPanelProps) {
  const updateFill = (
    update: (
      current: SpreadsheetFormatCellsFillDraft,
    ) => SpreadsheetFormatCellsFillDraft,
  ) => {
    touch('fill');
    setDraft((current) => ({
      ...current,
      fill: update(current.fill),
    }));
  };
  const activeFill = spreadsheetFormatCellsActiveFill(draft);

  return (
    <div className="work-spreadsheet-format-cells-fill">
      <SegmentedControl
        ariaLabel="填充类型"
        className="work-spreadsheet-format-cells-fill-modes"
        value={draft.fill.mode}
        items={fillModeItems}
        layout="equal"
        size="compact"
        onChange={(mode) => updateFill((current) => ({ ...current, mode }))}
      />

      <div className="work-spreadsheet-format-cells-fill-layout">
        <div className="work-spreadsheet-format-cells-fill-controls">
          {draft.fill.mode === 'none' && (
            <p className="work-spreadsheet-format-cells-fill-empty-copy">
              应用后会移除选区中的纯色、图案和渐变填充。
            </p>
          )}
          {draft.fill.mode === 'solid' && (
            <FillField label="背景色">
              <OfficeColorPicker
                ariaLabel="单元格填充颜色"
                value={draft.fill.solidColor}
                onValueChange={(solidColor) =>
                  updateFill((current) => ({ ...current, solidColor }))
                }
              />
            </FillField>
          )}
          {draft.fill.mode === 'pattern' && (
            <PatternFillControls
              fill={draft.fill.pattern}
              onChange={(pattern) =>
                updateFill((current) => ({ ...current, pattern }))
              }
            />
          )}
          {draft.fill.mode === 'gradient' && (
            <GradientFillControls
              fill={draft.fill.gradient}
              onChange={(gradient) =>
                updateFill((current) => ({ ...current, gradient }))
              }
            />
          )}
          {errors.fill && <p role="alert">{errors.fill}</p>}
          {source.fields.fill.mixed && !touched.fill && (
            <p className="work-spreadsheet-format-cells-mixed">
              选区包含多种填充；只有选择新的填充类型或参数后才会统一应用。
            </p>
          )}
        </div>
        <SpreadsheetFormatCellsFillPreview
          fill={activeFill}
          patternLabels={patternLabels}
        />
      </div>
    </div>
  );
}

function PatternFillControls({
  fill,
  onChange,
}: {
  fill: XlsxPatternFill;
  onChange: (fill: XlsxPatternFill) => void;
}) {
  const updateColor = (
    field: 'backgroundColor' | 'foregroundColor',
    color: string,
  ) => {
    const next = { ...fill, [field]: color };
    if (field === 'backgroundColor') delete next.backgroundColorOrigin;
    else delete next.foregroundColorOrigin;
    onChange(next);
  };
  return (
    <div className="work-spreadsheet-format-cells-pattern-controls">
      <FillField label="图案样式">
        <OfficeSelect
          ariaLabel="填充图案样式"
          value={fill.patternType}
          options={patternOptions}
          onValueChange={(patternType) => onChange({ ...fill, patternType })}
        />
      </FillField>
      <FillField label="图案颜色">
        <OfficeColorPicker
          ariaLabel="填充图案颜色"
          value={fill.foregroundColor}
          onValueChange={(color) => updateColor('foregroundColor', color)}
        />
      </FillField>
      <FillField label="背景色">
        <OfficeColorPicker
          ariaLabel="填充图案背景色"
          value={fill.backgroundColor}
          onValueChange={(color) => updateColor('backgroundColor', color)}
        />
      </FillField>
    </div>
  );
}

function GradientFillControls({
  fill,
  onChange,
}: {
  fill: XlsxGradientFill;
  onChange: (fill: XlsxGradientFill) => void;
}) {
  const changeType = (type: XlsxGradientFill['type']) => {
    if (type === fill.type) return;
    const stops = fill.stops.map((stop) => ({ ...stop }));
    onChange(
      type === 'linear'
        ? { degree: 0, stops, type }
        : {
            bottom: 0.75,
            left: 0.25,
            right: 0.75,
            stops,
            top: 0.25,
            type,
          },
    );
  };
  const changeStops = (stops: XlsxGradientStop[]) =>
    onChange({ ...fill, stops } as XlsxGradientFill);
  const updateStop = (
    index: number,
    update: (stop: XlsxGradientStop) => XlsxGradientStop,
  ) =>
    changeStops(
      fill.stops.map((stop, stopIndex) =>
        stopIndex === index ? update(stop) : stop,
      ),
    );

  return (
    <div
      className="work-spreadsheet-format-cells-gradient-controls"
      data-gradient-type={fill.type}
    >
      <div className="work-spreadsheet-format-cells-gradient-geometry">
        <FillField label="渐变类型">
          <OfficeSelect
            ariaLabel="渐变类型"
            value={fill.type}
            options={gradientTypeOptions}
            onValueChange={changeType}
          />
        </FillField>
        {fill.type === 'linear' ? (
          <PercentField
            label="角度（°）"
            ariaLabel="线性渐变角度"
            value={fill.degree}
            step={1}
            onChange={(degree) => onChange({ ...fill, degree })}
          />
        ) : (
          <PathGeometryFields fill={fill} onChange={onChange} />
        )}
      </div>

      <div className="work-spreadsheet-format-cells-gradient-stops-heading">
        <div>
          <strong>渐变色标</strong>
          <span>
            {fill.stops.length} / {MAX_XLSX_GRADIENT_STOPS}
          </span>
        </div>
        <Button
          size="compact"
          tone="quiet"
          disabled={fill.stops.length >= MAX_XLSX_GRADIENT_STOPS}
          onClick={() => changeStops(addGradientStop(fill.stops))}
        >
          <Plus size={13} aria-hidden="true" />
          添加色标
        </Button>
      </div>

      <ol
        className="work-spreadsheet-format-cells-gradient-stops"
        aria-label="渐变色标列表"
        data-stop-count={fill.stops.length}
      >
        {fill.stops.map((stop, index) => (
          <li
            className="work-spreadsheet-format-cells-gradient-stop"
            key={index}
          >
            <span className="work-spreadsheet-format-cells-gradient-stop-index">
              {index + 1}
            </span>
            <OfficeColorPicker
              ariaLabel={`色标 ${index + 1} 颜色`}
              value={stop.color}
              onValueChange={(color) =>
                updateStop(index, (current) => {
                  const next = { ...current, color };
                  delete next.colorOrigin;
                  return next;
                })
              }
            />
            <div className="work-spreadsheet-format-cells-stop-position">
              <span className="sr-only">色标 {index + 1} 位置</span>
              <OfficeNumberField
                ariaLabel={`色标 ${index + 1} 位置`}
                value={formatPercentage(stop.position)}
                min={0}
                max={100}
                step={1}
                validationInvalid={
                  stop.position < 0 ||
                  stop.position > 1 ||
                  (index > 0 &&
                    (fill.stops[index - 1]?.position ?? 0) > stop.position)
                }
                onValueChange={(value) => {
                  const percentage = Number(value);
                  if (!Number.isFinite(percentage)) return;
                  updateStop(index, (current) => ({
                    ...current,
                    position: percentage / 100,
                  }));
                }}
              />
              <span aria-hidden="true">%</span>
            </div>
            <IconButton
              className="work-spreadsheet-format-cells-gradient-stop-remove"
              label={`删除色标 ${index + 1}`}
              disabled={fill.stops.length <= 2}
              onClick={() =>
                changeStops(
                  fill.stops.filter((_, stopIndex) => stopIndex !== index),
                )
              }
            >
              <Trash2 size={13} />
            </IconButton>
          </li>
        ))}
      </ol>
      <small className="work-spreadsheet-format-cells-gradient-help">
        色标按列表顺序写入 XLSX；位置需从小到大排列。编辑颜色后会改用显式
        RGB，未修改的主题色语义会继续保留。
      </small>
    </div>
  );
}

function PathGeometryFields({
  fill,
  onChange,
}: {
  fill: Extract<XlsxGradientFill, { type: 'path' }>;
  onChange: (fill: XlsxGradientFill) => void;
}) {
  const fields = [
    ['left', '左边界'],
    ['right', '右边界'],
    ['top', '上边界'],
    ['bottom', '下边界'],
  ] as const;
  return (
    <div className="work-spreadsheet-format-cells-path-geometry">
      {fields.map(([field, label]) => (
        <PercentField
          key={field}
          label={`${label}（%）`}
          ariaLabel={`路径渐变${label}`}
          value={fill[field] * 100}
          min={0}
          max={100}
          step={1}
          invalid={
            (field === 'left' && fill.left > fill.right) ||
            (field === 'right' && fill.right < fill.left) ||
            (field === 'top' && fill.top > fill.bottom) ||
            (field === 'bottom' && fill.bottom < fill.top)
          }
          onChange={(value) => onChange({ ...fill, [field]: value / 100 })}
        />
      ))}
    </div>
  );
}

function PercentField({
  label,
  ariaLabel,
  value,
  min,
  max,
  step,
  invalid = false,
  onChange,
}: {
  label: string;
  ariaLabel: string;
  value: number;
  min?: number;
  max?: number;
  step: number;
  invalid?: boolean;
  onChange: (value: number) => void;
}) {
  return (
    <FillField label={label}>
      <OfficeNumberField
        ariaLabel={ariaLabel}
        value={formatNumber(value)}
        min={min}
        max={max}
        step={step}
        validationInvalid={invalid}
        onValueChange={(rawValue) => {
          const next = Number(rawValue);
          if (Number.isFinite(next)) onChange(next);
        }}
      />
    </FillField>
  );
}

function FillField({
  label,
  children,
}: {
  label: string;
  children: ReactNode;
}) {
  return (
    <div className="work-spreadsheet-format-cells-field">
      <span>{label}</span>
      {children}
    </div>
  );
}

function addGradientStop(
  source: readonly XlsxGradientStop[],
): XlsxGradientStop[] {
  if (source.length >= MAX_XLSX_GRADIENT_STOPS) return [...source];
  let insertAfter = 0;
  let largestGap = Number.NEGATIVE_INFINITY;
  for (let index = 0; index < source.length - 1; index += 1) {
    const left = source[index];
    const right = source[index + 1];
    if (!left || !right) continue;
    const gap = right.position - left.position;
    if (gap > largestGap) {
      insertAfter = index;
      largestGap = gap;
    }
  }
  const left = source[insertAfter] ?? { color: '#4472c4', position: 0 };
  const right = source[insertAfter + 1] ?? {
    color: '#ffffff',
    position: 1,
  };
  const stop = {
    color: interpolateHexColor(left.color, right.color),
    position: (left.position + right.position) / 2,
  };
  return [
    ...source.slice(0, insertAfter + 1).map((item) => ({ ...item })),
    stop,
    ...source.slice(insertAfter + 1).map((item) => ({ ...item })),
  ];
}

function interpolateHexColor(left: string, right: string): string {
  const channels = [1, 3, 5].map((offset) =>
    Math.round(
      (Number.parseInt(left.slice(offset, offset + 2), 16) +
        Number.parseInt(right.slice(offset, offset + 2), 16)) /
        2,
    ),
  );
  return `#${channels
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')}`;
}

function formatPercentage(value: number): string {
  return formatNumber(value * 100);
}

function formatNumber(value: number): string {
  return Number.isInteger(value)
    ? String(value)
    : String(Number(value.toFixed(4)));
}
