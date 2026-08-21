import type { Dispatch, ReactNode, SetStateAction } from 'react';
import { Button } from '../../../design-system/primitives';
import type {
  SpreadsheetCellBorderFormat,
  SpreadsheetCellBorderStyle,
} from './spreadsheet-cell-border';
import type { SpreadsheetUnderlineStyle } from '../work-spreadsheet-underline';
import type { SpreadsheetCellFormatPatch } from './spreadsheet-cell-format';
import type {
  SpreadsheetFormatCellsDialogSource,
  SpreadsheetFormatCellsDraft,
  SpreadsheetFormatCellsDraftErrors,
  SpreadsheetFormatCellsTabId,
  SpreadsheetFormatCellsTouched,
} from './spreadsheet-format-cells-dialog-model';
import {
  spreadsheetFontFamilyOptions,
  spreadsheetFontSizeOptions,
} from './spreadsheet-editor-support';
import {
  type SpreadsheetNumberFormatPreset,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatPreset,
  spreadsheetNumberFormatPresetLabels,
} from './spreadsheet-number-format';
import {
  OfficeCheckbox,
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
  type OfficeSelectOption,
  OfficeTextField,
} from './office-controls';

const MIXED_VALUE = '__mixed__';

interface PanelProps {
  activeTab: SpreadsheetFormatCellsTabId;
  idBase: string;
  source: SpreadsheetFormatCellsDialogSource;
  draft: SpreadsheetFormatCellsDraft;
  errors: SpreadsheetFormatCellsDraftErrors;
  touched: SpreadsheetFormatCellsTouched;
  setDraft: Dispatch<SetStateAction<SpreadsheetFormatCellsDraft>>;
  touch: (field: keyof SpreadsheetCellFormatPatch) => void;
}

const numberFormatOptions = [
  'general',
  'number',
  'currency',
  'accounting',
  'percent',
  'date',
  'time',
  'scientific',
  'fraction',
  'text',
  'custom',
].map((value) => ({
  value: value as SpreadsheetNumberFormatPreset,
  label:
    spreadsheetNumberFormatPresetLabels[value as SpreadsheetNumberFormatPreset],
}));

const horizontalOptions = [
  { value: 'general', label: '常规' },
  { value: 'left', label: '左对齐' },
  { value: 'center', label: '居中' },
  { value: 'right', label: '右对齐' },
] as const;

const verticalOptions = [
  { value: 'top', label: '顶端对齐' },
  { value: 'middle', label: '垂直居中' },
  { value: 'bottom', label: '底端对齐' },
] as const;

const underlineOptions = [
  { value: 'none', label: '无' },
  { value: 'single', label: '单下划线' },
  { value: 'double', label: '双下划线' },
  { value: 'singleAccounting', label: '单会计用下划线' },
  { value: 'doubleAccounting', label: '双会计用下划线' },
] as const satisfies readonly OfficeSelectOption<SpreadsheetUnderlineStyle>[];

const borderStyleOptions = [
  { value: 'thin', label: '细实线' },
  { value: 'dotted', label: '点线' },
  { value: 'dashed', label: '虚线' },
  { value: 'dash-dot', label: '点划线' },
  { value: 'medium', label: '中等实线' },
  { value: 'medium-dashed', label: '中等虚线' },
  { value: 'thick', label: '粗实线' },
] as const satisfies readonly OfficeSelectOption<SpreadsheetCellBorderStyle>[];

const borderTargets = [
  { target: 'top', label: '上框线' },
  { target: 'bottom', label: '下框线' },
  { target: 'left', label: '左框线' },
  { target: 'right', label: '右框线' },
  { target: 'diagonal', label: '对角线' },
] as const;

export function SpreadsheetFormatCellsPanel(props: PanelProps) {
  return (
    <section
      id={`${props.idBase}-${props.activeTab}-panel`}
      className="work-spreadsheet-format-cells-panel"
      role="tabpanel"
      aria-labelledby={`${props.idBase}-${props.activeTab}-tab`}
      tabIndex={-1}
    >
      {props.activeTab === 'number' && <NumberPanel {...props} />}
      {props.activeTab === 'alignment' && <AlignmentPanel {...props} />}
      {props.activeTab === 'font' && <FontPanel {...props} />}
      {props.activeTab === 'border' && <BorderPanel {...props} />}
      {props.activeTab === 'fill' && <FillPanel {...props} />}
      {props.activeTab === 'protection' && <ProtectionPanel {...props} />}
    </section>
  );
}

function NumberPanel({
  source,
  draft,
  errors,
  touched,
  setDraft,
  touch,
}: PanelProps) {
  const mixed = source.fields.numberFormat.mixed && !touched.numberFormat;
  const preset = spreadsheetNumberFormatPreset(draft.numberFormat);
  const options = withMixedOption(numberFormatOptions, mixed);
  return (
    <div className="work-spreadsheet-format-cells-number">
      <Field label="分类">
        <OfficeSelect
          ariaLabel="数字格式分类"
          value={mixed ? MIXED_VALUE : preset}
          options={options}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('numberFormat');
            if (value !== 'custom') {
              setDraft((current) => ({
                ...current,
                numberFormat: spreadsheetNumberFormatCode(value),
              }));
            }
          }}
        />
      </Field>
      <Field label="格式代码">
        <OfficeTextField
          aria-label="数字格式代码"
          aria-invalid={Boolean(errors.numberFormat) || undefined}
          value={draft.numberFormat}
          onChange={(event) => {
            const numberFormat = event.currentTarget.value;
            touch('numberFormat');
            setDraft((current) => ({
              ...current,
              numberFormat,
            }));
          }}
        />
      </Field>
      {errors.numberFormat && <p role="alert">{errors.numberFormat}</p>}
      <div className="work-spreadsheet-format-cells-sample">
        <span>示例</span>
        <output aria-label="数字格式示例">
          {String(source.activeCell?.m ?? source.activeCell?.v ?? '1234.56')}
        </output>
        <small>{draft.numberFormat || '—'}</small>
      </div>
      {mixed && <MixedHint />}
    </div>
  );
}

function AlignmentPanel({
  source,
  draft,
  errors,
  touched,
  setDraft,
  touch,
}: PanelProps) {
  const horizontalMixed =
    source.fields.horizontalAlignment.mixed && !touched.horizontalAlignment;
  const verticalMixed =
    source.fields.verticalAlignment.mixed && !touched.verticalAlignment;
  return (
    <div className="work-spreadsheet-format-cells-grid">
      <Field label="水平对齐">
        <OfficeSelect
          ariaLabel="水平对齐"
          value={horizontalMixed ? MIXED_VALUE : draft.horizontalAlignment}
          options={withMixedOption(horizontalOptions, horizontalMixed)}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('horizontalAlignment');
            setDraft((current) => ({ ...current, horizontalAlignment: value }));
          }}
        />
      </Field>
      <Field label="垂直对齐">
        <OfficeSelect
          ariaLabel="垂直对齐"
          value={verticalMixed ? MIXED_VALUE : draft.verticalAlignment}
          options={withMixedOption(verticalOptions, verticalMixed)}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('verticalAlignment');
            setDraft((current) => ({ ...current, verticalAlignment: value }));
          }}
        />
      </Field>
      <fieldset className="work-spreadsheet-format-cells-options">
        <legend>文字控制</legend>
        <OfficeCheckbox
          ariaLabel="自动换行"
          checked={draft.wrapText}
          indeterminate={source.fields.wrapText.mixed && !touched.wrapText}
          onCheckedChange={(checked) => {
            touch('wrapText');
            setDraft((current) => ({ ...current, wrapText: checked }));
          }}
        >
          自动换行
        </OfficeCheckbox>
      </fieldset>
      <Field label="文字旋转（度）">
        <OfficeNumberField
          ariaLabel="文字旋转角度"
          value={draft.rotation}
          min={-90}
          max={90}
          validationInvalid={Boolean(errors.rotation)}
          onValueChange={(value) => {
            const rotation = Number(value);
            if (!Number.isFinite(rotation)) return;
            touch('rotation');
            setDraft((current) => ({ ...current, rotation }));
          }}
        />
      </Field>
      {errors.rotation && <p role="alert">{errors.rotation}</p>}
      {(horizontalMixed || verticalMixed || source.fields.rotation.mixed) && (
        <MixedHint />
      )}
    </div>
  );
}

function FontPanel({
  source,
  draft,
  errors,
  touched,
  setDraft,
  touch,
}: PanelProps) {
  const familyMixed = source.fields.fontFamily.mixed && !touched.fontFamily;
  const sizeMixed = source.fields.fontSize.mixed && !touched.fontSize;
  const underlineMixed = source.fields.underline.mixed && !touched.underline;
  const toggle = (field: 'bold' | 'italic' | 'strike', checked: boolean) => {
    touch(field);
    setDraft((current) => ({ ...current, [field]: checked }));
  };
  return (
    <div className="work-spreadsheet-format-cells-grid font">
      <Field label="字体">
        <OfficeSelect
          ariaLabel="单元格字体"
          value={familyMixed ? MIXED_VALUE : draft.fontFamily}
          options={withMixedOption(
            spreadsheetFontFamilyOptions(draft.fontFamily),
            familyMixed,
          )}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('fontFamily');
            setDraft((current) => ({ ...current, fontFamily: value }));
          }}
        />
      </Field>
      <Field label="字号">
        <OfficeSelect
          ariaLabel="单元格字号"
          value={sizeMixed ? MIXED_VALUE : String(draft.fontSize)}
          options={withMixedOption(
            spreadsheetFontSizeOptions(draft.fontSize),
            sizeMixed,
          )}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('fontSize');
            setDraft((current) => ({ ...current, fontSize: Number(value) }));
          }}
        />
      </Field>
      {errors.fontSize && <p role="alert">{errors.fontSize}</p>}
      <fieldset className="work-spreadsheet-format-cells-options emphasis">
        <legend>字形</legend>
        {(
          [
            ['bold', '加粗'],
            ['italic', '斜体'],
            ['strike', '删除线'],
          ] as const
        ).map(([field, label]) => (
          <OfficeCheckbox
            key={field}
            ariaLabel={label}
            checked={draft[field]}
            indeterminate={source.fields[field].mixed && !touched[field]}
            onCheckedChange={(checked) => toggle(field, checked)}
          >
            {label}
          </OfficeCheckbox>
        ))}
      </fieldset>
      <Field label="下划线">
        <OfficeSelect
          ariaLabel="下划线样式"
          value={underlineMixed ? MIXED_VALUE : draft.underline}
          options={withMixedOption(underlineOptions, underlineMixed)}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('underline');
            setDraft((current) => ({ ...current, underline: value }));
          }}
        />
      </Field>
      <Field label="文字颜色">
        <OfficeColorPicker
          ariaLabel="单元格文字颜色"
          value={draft.fontColor}
          onValueChange={(fontColor) => {
            touch('fontColor');
            setDraft((current) => ({ ...current, fontColor }));
          }}
        />
      </Field>
      <div
        className="work-spreadsheet-format-cells-font-preview"
        style={{
          color: draft.fontColor,
          fontFamily: draft.fontFamily,
          fontSize: `${Math.max(10, draft.fontSize)}px`,
          fontStyle: draft.italic ? 'italic' : 'normal',
          fontWeight: draft.bold ? 700 : 400,
          textDecorationLine: [
            draft.underline !== 'none' && 'underline',
            draft.strike && 'line-through',
          ]
            .filter(Boolean)
            .join(' '),
          textDecorationStyle:
            draft.underline === 'double' ||
            draft.underline === 'doubleAccounting'
              ? 'double'
              : 'solid',
          textUnderlineOffset:
            draft.underline === 'singleAccounting' ||
            draft.underline === 'doubleAccounting'
              ? '0.24em'
              : '0.12em',
        }}
      >
        A3S Office 字体预览
      </div>
    </div>
  );
}

function BorderPanel({ source, draft, touched, setDraft, touch }: PanelProps) {
  const hasTarget = (target: SpreadsheetCellBorderFormat['target']) =>
    draft.borders.some((format) => format.target === target);
  const updateBorders = (borders: SpreadsheetCellBorderFormat[]) => {
    touch('borders');
    setDraft((current) => ({ ...current, borders }));
  };
  const restyleBorders = (
    field: 'borderColor' | 'borderStyle',
    value: string,
  ) => {
    setDraft((current) => {
      const next = {
        ...current,
        [field]: value,
      } as SpreadsheetFormatCellsDraft;
      if (current.borders.length) {
        touch('borders');
        next.borders = current.borders.map((format) => ({
          ...format,
          color: field === 'borderColor' ? value : current.borderColor,
          style:
            field === 'borderStyle'
              ? (value as SpreadsheetCellBorderStyle)
              : current.borderStyle,
        }));
      }
      return next;
    });
  };
  return (
    <div className="work-spreadsheet-format-cells-border">
      <div className="work-spreadsheet-format-cells-border-tools">
        <Field label="线条样式">
          <OfficeSelect
            ariaLabel="边框线条样式"
            value={draft.borderStyle}
            options={borderStyleOptions}
            onValueChange={(value) => restyleBorders('borderStyle', value)}
          />
        </Field>
        <Field label="线条颜色">
          <OfficeColorPicker
            ariaLabel="边框线条颜色"
            value={draft.borderColor}
            onValueChange={(value) => restyleBorders('borderColor', value)}
          />
        </Field>
      </div>
      <div className="work-spreadsheet-format-cells-border-layout">
        <div className="work-spreadsheet-format-cells-border-actions">
          {borderTargets.map(({ target, label }) => (
            <button
              key={target}
              type="button"
              className={`work-spreadsheet-border-edge ${target}`}
              aria-label={label}
              aria-pressed={hasTarget(target)}
              onClick={() =>
                updateBorders(
                  hasTarget(target)
                    ? draft.borders.filter((format) => format.target !== target)
                    : [
                        ...draft.borders,
                        {
                          target,
                          color: draft.borderColor,
                          style: draft.borderStyle,
                        },
                      ],
                )
              }
            >
              <span aria-hidden="true" />
            </button>
          ))}
          <Button size="compact" tone="quiet" onClick={() => updateBorders([])}>
            无边框
          </Button>
        </div>
        <BorderPreview draft={draft} />
      </div>
      {source.fields.borders.mixed && !touched.borders && <MixedHint />}
    </div>
  );
}

function BorderPreview({ draft }: { draft: SpreadsheetFormatCellsDraft }) {
  const line = (target: SpreadsheetCellBorderFormat['target']) => {
    const format = draft.borders.find(
      (candidate) => candidate.target === target,
    );
    return format ? borderCss(format) : undefined;
  };
  return (
    <div
      className="work-spreadsheet-format-cells-border-preview"
      role="img"
      aria-label="边框预览"
    >
      <span className="top" style={{ borderTop: line('top') }} />
      <span className="bottom" style={{ borderBottom: line('bottom') }} />
      <span className="left" style={{ borderLeft: line('left') }} />
      <span className="right" style={{ borderRight: line('right') }} />
      {line('diagonal') && (
        <span className="diagonal" style={{ borderTop: line('diagonal') }} />
      )}
      <strong>文本</strong>
    </div>
  );
}

function FillPanel({ source, draft, touched, setDraft, touch }: PanelProps) {
  return (
    <div className="work-spreadsheet-format-cells-fill">
      <Field label="背景色">
        <OfficeColorPicker
          ariaLabel="单元格填充颜色"
          value={draft.fillColor ?? '#ffffff'}
          onValueChange={(fillColor) => {
            touch('fillColor');
            setDraft((current) => ({ ...current, fillColor }));
          }}
        />
      </Field>
      <Button
        size="compact"
        tone="quiet"
        onClick={() => {
          touch('fillColor');
          setDraft((current) => ({ ...current, fillColor: null }));
        }}
      >
        无填充
      </Button>
      <div
        className={`work-spreadsheet-format-cells-fill-preview${draft.fillColor ? '' : ' empty'}`}
        style={{ backgroundColor: draft.fillColor ?? undefined }}
      >
        {draft.fillColor ? draft.fillColor.toUpperCase() : '无填充'}
      </div>
      {source.fields.fillColor.mixed && !touched.fillColor && <MixedHint />}
    </div>
  );
}

function ProtectionPanel({
  source,
  draft,
  touched,
  setDraft,
  touch,
}: PanelProps) {
  return (
    <div className="work-spreadsheet-format-cells-protection">
      <OfficeCheckbox
        ariaLabel="锁定单元格"
        checked={draft.locked}
        indeterminate={source.fields.locked.mixed && !touched.locked}
        onCheckedChange={(locked) => {
          touch('locked');
          setDraft((current) => ({ ...current, locked }));
        }}
      >
        锁定单元格
      </OfficeCheckbox>
      <OfficeCheckbox
        ariaLabel="隐藏公式"
        checked={draft.hidden}
        indeterminate={source.fields.hidden.mixed && !touched.hidden}
        onCheckedChange={(hidden) => {
          touch('hidden');
          setDraft((current) => ({ ...current, hidden }));
        }}
      >
        隐藏公式
      </OfficeCheckbox>
      <p>保护属性会随“应用”一起保存；锁定和隐藏效果在工作表保护启用后生效。</p>
    </div>
  );
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="work-spreadsheet-format-cells-field">
      <span>{label}</span>
      {children}
    </div>
  );
}

function MixedHint() {
  return (
    <p className="work-spreadsheet-format-cells-mixed">
      选区包含多种设置；仅修改的项目会统一应用。
    </p>
  );
}

function withMixedOption<T extends string>(
  options: readonly OfficeSelectOption<T>[],
  mixed: boolean,
): readonly OfficeSelectOption<T | typeof MIXED_VALUE>[] {
  return mixed
    ? [{ value: MIXED_VALUE, label: '混合', disabled: true }, ...options]
    : options;
}

function borderCss(format: SpreadsheetCellBorderFormat): string {
  const width =
    format.style === 'thick' ? 3 : format.style.startsWith('medium') ? 2 : 1;
  const style =
    format.style === 'dotted'
      ? 'dotted'
      : format.style.includes('dash')
        ? 'dashed'
        : 'solid';
  return `${width}px ${style} ${format.color}`;
}
