import { Plus, Trash2 } from 'lucide-react';
import type { Dispatch, ReactNode, SetStateAction } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
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
import { useOfficeMessages } from './office-messages-context';

interface FillPanelProps {
  source: SpreadsheetFormatCellsDialogSource;
  draft: SpreadsheetFormatCellsDraft;
  errors: SpreadsheetFormatCellsDraftErrors;
  touched: SpreadsheetFormatCellsTouched;
  setDraft: Dispatch<SetStateAction<SpreadsheetFormatCellsDraft>>;
  touch: (field: keyof SpreadsheetCellFormatPatch) => void;
}

function fillModeItems(messages: OfficeMessageCatalog) {
  return [
    { id: 'none' as const, label: officeMessage(messages, 'spreadsheet.formatCells.fill.mode.none') },
    { id: 'solid' as const, label: officeMessage(messages, 'spreadsheet.formatCells.fill.mode.solid') },
    { id: 'pattern' as const, label: officeMessage(messages, 'spreadsheet.formatCells.fill.mode.pattern') },
    { id: 'gradient' as const, label: officeMessage(messages, 'spreadsheet.formatCells.fill.mode.gradient') },
  ];
}

function gradientTypeOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<XlsxGradientFill['type']>[] {
  return [
    { value: 'linear', label: officeMessage(messages, 'spreadsheet.formatCells.fill.gradient.linear') },
    { value: 'path', label: officeMessage(messages, 'spreadsheet.formatCells.fill.gradient.path') },
  ];
}

function patternLabels(
  messages: OfficeMessageCatalog,
): Record<XlsxPatternFillType, string> {
  return {
    darkDown: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.darkDown'),
    darkGray: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.darkGray'),
    darkGrid: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.darkGrid'),
    darkHorizontal: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.darkHorizontal'),
    darkTrellis: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.darkTrellis'),
    darkUp: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.darkUp'),
    darkVertical: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.darkVertical'),
    gray0625: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.gray0625'),
    gray125: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.gray125'),
    lightDown: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.lightDown'),
    lightGray: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.lightGray'),
    lightGrid: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.lightGrid'),
    lightHorizontal: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.lightHorizontal'),
    lightTrellis: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.lightTrellis'),
    lightUp: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.lightUp'),
    lightVertical: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.lightVertical'),
    mediumGray: officeMessage(messages, 'spreadsheet.formatCells.fill.pattern.mediumGray'),
  };
}

function patternOptions(messages: OfficeMessageCatalog) {
  const labels = patternLabels(messages);
  return xlsxPatternFillTypes.map((value) => ({
    label: labels[value],
    value,
  }));
}

export function SpreadsheetFormatCellsFillPanel({
  source,
  draft,
  errors,
  touched,
  setDraft,
  touch,
}: FillPanelProps) {
  const messages = useOfficeMessages();
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
  const fillMixed = source.fields.fill.mixed && !touched.fill;
  const activeFill = spreadsheetFormatCellsActiveFill(draft);

  return (
    <div className="work-spreadsheet-format-cells-fill">
      <SegmentedControl
        ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.modeAria')}
        className="work-spreadsheet-format-cells-fill-modes"
        value={fillMixed ? null : draft.fill.mode}
        items={fillModeItems(messages)}
        layout="equal"
        size="compact"
        onChange={(mode) => updateFill((current) => ({ ...current, mode }))}
      />

      <div className="work-spreadsheet-format-cells-fill-layout">
        <div className="work-spreadsheet-format-cells-fill-controls">
          {draft.fill.mode === 'none' && (
            <p className="work-spreadsheet-format-cells-fill-empty-copy">
              {officeMessage(messages, 'spreadsheet.formatCells.fill.noneHint')}
            </p>
          )}
          {draft.fill.mode === 'solid' && (
            <FillField label={officeMessage(messages, 'spreadsheet.formatCells.fill.bgColor')}>
              <OfficeColorPicker
                ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.solidColorAria')}
                value={draft.fill.solidColor}
                triggerLabel={
                  fillMixed
                    ? officeMessage(messages, 'spreadsheet.formatCells.mixed')
                    : undefined
                }
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
              {officeMessage(messages, 'spreadsheet.formatCells.fill.mixedHint')}
            </p>
          )}
        </div>
        <SpreadsheetFormatCellsFillPreview
          fill={activeFill}
          patternLabels={patternLabels(messages)}
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
  const messages = useOfficeMessages();
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
      <FillField label={officeMessage(messages, 'spreadsheet.formatCells.fill.patternStyle')}>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.patternStyleAria')}
          value={fill.patternType}
          options={patternOptions(messages)}
          onValueChange={(patternType) => onChange({ ...fill, patternType })}
        />
      </FillField>
      <FillField label={officeMessage(messages, 'spreadsheet.formatCells.fill.patternColor')}>
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.patternColorAria')}
          value={fill.foregroundColor}
          onValueChange={(color) => updateColor('foregroundColor', color)}
        />
      </FillField>
      <FillField label={officeMessage(messages, 'spreadsheet.formatCells.fill.bgColor')}>
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.patternBgAria')}
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
  const messages = useOfficeMessages();
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
        <FillField label={officeMessage(messages, 'spreadsheet.formatCells.fill.gradientType')}>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.gradientType')}
            value={fill.type}
            options={gradientTypeOptions(messages)}
            onValueChange={changeType}
          />
        </FillField>
        {fill.type === 'linear' ? (
          <PercentField
            label={officeMessage(messages, 'spreadsheet.formatCells.fill.angle')}
            ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.angleAria')}
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
          <strong>{officeMessage(messages, 'spreadsheet.formatCells.fill.stops')}</strong>
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
          {officeMessage(messages, 'spreadsheet.formatCells.fill.addStop')}
        </Button>
      </div>

      <ol
        className="work-spreadsheet-format-cells-gradient-stops"
        aria-label={officeMessage(messages, 'spreadsheet.formatCells.fill.stopsListAria')}
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
              ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.stopColorAria', { n: String(index + 1) })}
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
              <span className="sr-only">{officeMessage(messages, 'spreadsheet.formatCells.fill.stopPosSr', { n: String(index + 1) })}</span>
              <OfficeNumberField
                ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.stopPosAria', { n: String(index + 1) })}
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
              label={officeMessage(messages, 'spreadsheet.formatCells.fill.deleteStop', { n: String(index + 1) })}
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
        {officeMessage(messages, 'spreadsheet.formatCells.fill.stopsNote')}
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
  const messages = useOfficeMessages();
  const fields = [
    ['left', officeMessage(messages, 'spreadsheet.formatCells.fill.path.left')],
    ['right', officeMessage(messages, 'spreadsheet.formatCells.fill.path.right')],
    ['top', officeMessage(messages, 'spreadsheet.formatCells.fill.path.top')],
    ['bottom', officeMessage(messages, 'spreadsheet.formatCells.fill.path.bottom')],
  ] as const;
  return (
    <div className="work-spreadsheet-format-cells-path-geometry">
      {fields.map(([field, label]) => (
        <PercentField
          key={field}
          label={officeMessage(messages, 'spreadsheet.formatCells.fill.pathBoundPercent', { label })}
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.fill.pathBoundAria', { label })}
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
