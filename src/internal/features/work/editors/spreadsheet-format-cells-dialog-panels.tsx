import type { CSSProperties, Dispatch, ReactNode, SetStateAction } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
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
import type { SpreadsheetFormatCellsInitialFocus } from './spreadsheet-format-cells-intent';
import { SpreadsheetFormatCellsFillPanel } from './spreadsheet-format-cells-fill-panel';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import {
  spreadsheetFontFamilyOptions,
  spreadsheetFontSizeOptions,
} from './spreadsheet-editor-support';
import {
  type SpreadsheetNumberFormatPreset,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatPreset,
  spreadsheetNumberFormatPresetLabels,
  spreadsheetNumberFormatPreview,
} from './spreadsheet-number-format';
import { useOfficeMessages } from './office-messages-context';
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
  initialFocus?: SpreadsheetFormatCellsInitialFocus;
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

function horizontalOptions(messages: OfficeMessageCatalog) {
  return [
    { value: 'general' as const, label: officeMessage(messages, 'spreadsheet.formatCells.align.horizontal.general') },
    { value: 'left' as const, label: officeMessage(messages, 'spreadsheet.formatCells.align.horizontal.left') },
    { value: 'center' as const, label: officeMessage(messages, 'spreadsheet.formatCells.align.horizontal.center') },
    { value: 'right' as const, label: officeMessage(messages, 'spreadsheet.formatCells.align.horizontal.right') },
  ];
}

function verticalOptions(messages: OfficeMessageCatalog) {
  return [
    { value: 'top' as const, label: officeMessage(messages, 'spreadsheet.formatCells.align.vertical.top') },
    { value: 'middle' as const, label: officeMessage(messages, 'spreadsheet.formatCells.align.vertical.middle') },
    { value: 'bottom' as const, label: officeMessage(messages, 'spreadsheet.formatCells.align.vertical.bottom') },
  ];
}

function underlineOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<SpreadsheetUnderlineStyle>[] {
  return [
    { value: 'none', label: officeMessage(messages, 'spreadsheet.formatCells.underline.none') },
    { value: 'single', label: officeMessage(messages, 'spreadsheet.formatCells.underline.single') },
    { value: 'double', label: officeMessage(messages, 'spreadsheet.formatCells.underline.double') },
    { value: 'singleAccounting', label: officeMessage(messages, 'spreadsheet.formatCells.underline.singleAccounting') },
    { value: 'doubleAccounting', label: officeMessage(messages, 'spreadsheet.formatCells.underline.doubleAccounting') },
  ];
}

function borderStyleOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<SpreadsheetCellBorderStyle>[] {
  return [
    { value: 'thin', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.thin') },
    { value: 'dotted', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.dotted') },
    { value: 'dashed', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.dashed') },
    { value: 'dash-dot', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.dashDot') },
    { value: 'dash-dot-dot', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.dashDotDot') },
    { value: 'medium', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.medium') },
    { value: 'medium-dashed', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.mediumDashed') },
    { value: 'medium-dash-dot', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.mediumDashDot') },
    { value: 'medium-dash-dot-dot', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.mediumDashDotDot') },
    { value: 'thick', label: officeMessage(messages, 'spreadsheet.formatCells.border.style.thick') },
  ];
}

function borderTargets(messages: OfficeMessageCatalog) {
  return [
    { target: 'top' as const, label: officeMessage(messages, 'spreadsheet.formatCells.border.target.top') },
    { target: 'bottom' as const, label: officeMessage(messages, 'spreadsheet.formatCells.border.target.bottom') },
    { target: 'left' as const, label: officeMessage(messages, 'spreadsheet.formatCells.border.target.left') },
    { target: 'right' as const, label: officeMessage(messages, 'spreadsheet.formatCells.border.target.right') },
    { target: 'diagonalDown' as const, label: officeMessage(messages, 'spreadsheet.formatCells.border.target.diagonalDown') },
    { target: 'diagonalUp' as const, label: officeMessage(messages, 'spreadsheet.formatCells.border.target.diagonalUp') },
  ];
}

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
      {props.activeTab === 'fill' && (
        <SpreadsheetFormatCellsFillPanel {...props} />
      )}
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
  const messages = useOfficeMessages();
  const mixed = source.fields.numberFormat.mixed && !touched.numberFormat;
  const preset = spreadsheetNumberFormatPreset(draft.numberFormat);
  const options = withMixedOption(numberFormatOptions, mixed, messages);
  return (
    <div className="work-spreadsheet-format-cells-number">
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.number.category')}>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.number.categoryAria')}
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
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.number.code')}>
        <OfficeTextField
          aria-label={officeMessage(messages, 'spreadsheet.formatCells.number.codeAria')}
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
        <span>{officeMessage(messages, 'spreadsheet.formatCells.number.sample')}</span>
        <output aria-label={officeMessage(messages, 'spreadsheet.formatCells.number.sampleAria')}>
          {spreadsheetNumberFormatPreview(draft.numberFormat, source.activeCell)}
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
  const messages = useOfficeMessages();
  const horizontalMixed =
    source.fields.horizontalAlignment.mixed && !touched.horizontalAlignment;
  const verticalMixed =
    source.fields.verticalAlignment.mixed && !touched.verticalAlignment;
  const rotationMixed = source.fields.rotation.mixed && !touched.rotation;
  return (
    <div className="work-spreadsheet-format-cells-grid">
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.align.horizontal')}>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.align.horizontal')}
          value={horizontalMixed ? MIXED_VALUE : draft.horizontalAlignment}
          options={withMixedOption(horizontalOptions(messages), horizontalMixed, messages)}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('horizontalAlignment');
            setDraft((current) => ({ ...current, horizontalAlignment: value }));
          }}
        />
      </Field>
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.align.vertical')}>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.align.vertical')}
          value={verticalMixed ? MIXED_VALUE : draft.verticalAlignment}
          options={withMixedOption(verticalOptions(messages), verticalMixed, messages)}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('verticalAlignment');
            setDraft((current) => ({ ...current, verticalAlignment: value }));
          }}
        />
      </Field>
      <fieldset className="work-spreadsheet-format-cells-options">
        <legend>{officeMessage(messages, 'spreadsheet.formatCells.align.textControl')}</legend>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.align.wrap')}
          checked={draft.wrapText}
          indeterminate={source.fields.wrapText.mixed && !touched.wrapText}
          onCheckedChange={(checked) => {
            touch('wrapText');
            setDraft((current) => ({ ...current, wrapText: checked }));
          }}
        >
          {officeMessage(messages, 'spreadsheet.formatCells.align.wrap')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.align.stacked')}
          checked={draft.stackedText}
          indeterminate={source.fields.stackedText.mixed && !touched.textOrientation}
          onCheckedChange={(checked) => {
            touch('textOrientation');
            setDraft((current) => ({ ...current, stackedText: checked }));
          }}
        >
          {officeMessage(messages, 'spreadsheet.formatCells.align.stacked')}
        </OfficeCheckbox>
      </fieldset>
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.align.rotation')}>
        <OfficeNumberField
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.align.rotationAria')}
          value={rotationMixed ? '' : draft.rotation}
          min={-90}
          max={90}
          disabled={draft.stackedText}
          placeholder={
            rotationMixed
              ? officeMessage(messages, 'spreadsheet.formatCells.mixed')
              : undefined
          }
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
  initialFocus,
  setDraft,
  touch,
}: PanelProps) {
  const messages = useOfficeMessages();
  const familyMixed = source.fields.fontFamily.mixed && !touched.fontFamily;
  const sizeMixed = source.fields.fontSize.mixed && !touched.fontSize;
  const underlineMixed = source.fields.underline.mixed && !touched.underline;
  const colorMixed = source.fields.fontColor.mixed && !touched.fontColor;
  const strikeMixed = source.fields.strike.mixed && !touched.strike;
  const toggle = (field: 'bold' | 'italic' | 'strike', checked: boolean) => {
    touch(field);
    setDraft((current) => ({ ...current, [field]: checked }));
  };
  return (
    <div className="work-spreadsheet-format-cells-grid font">
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.font.family')}>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.font.familyAria')}
          ariaKeyShortcuts={
            spreadsheetCommandCatalog.formatCellsFont.shortcut.aria
          }
          initialFocus={initialFocus === 'fontFamily'}
          value={familyMixed ? MIXED_VALUE : draft.fontFamily}
          options={withMixedOption(
            spreadsheetFontFamilyOptions(draft.fontFamily),
            familyMixed,
            messages,
          )}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('fontFamily');
            setDraft((current) => ({ ...current, fontFamily: value }));
          }}
        />
      </Field>
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.font.size')}>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.font.sizeAria')}
          ariaKeyShortcuts={
            spreadsheetCommandCatalog.formatCellsFontSize.shortcut.aria
          }
          initialFocus={initialFocus === 'fontSize'}
          value={sizeMixed ? MIXED_VALUE : String(draft.fontSize)}
          options={withMixedOption(
            spreadsheetFontSizeOptions(draft.fontSize),
            sizeMixed,
            messages,
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
        <legend>{officeMessage(messages, 'spreadsheet.formatCells.font.style')}</legend>
        {(
          [
            ['bold', officeMessage(messages, 'spreadsheet.formatCells.font.bold')],
            ['italic', officeMessage(messages, 'spreadsheet.formatCells.font.italic')],
            ['strike', officeMessage(messages, 'spreadsheet.formatCells.font.strike')],
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
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.font.underline')}>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.font.underlineAria')}
          value={underlineMixed ? MIXED_VALUE : draft.underline}
          options={withMixedOption(underlineOptions(messages), underlineMixed, messages)}
          onValueChange={(value) => {
            if (value === MIXED_VALUE) return;
            touch('underline');
            setDraft((current) => ({ ...current, underline: value }));
          }}
        />
      </Field>
      <Field label={officeMessage(messages, 'spreadsheet.formatCells.font.color')}>
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.font.colorAria')}
          value={draft.fontColor}
          triggerLabel={
            colorMixed
              ? officeMessage(messages, 'spreadsheet.formatCells.mixed')
              : undefined
          }
          onValueChange={(fontColor) => {
            touch('fontColor');
            setDraft((current) => ({ ...current, fontColor }));
          }}
        />
      </Field>
      <div
        className="work-spreadsheet-format-cells-font-preview"
        data-underline-style={underlineMixed ? undefined : draft.underline}
        data-strike={strikeMixed || !draft.strike ? undefined : 'true'}
        style={{
          backgroundColor: spreadsheetFontPreviewBackground(draft.fontColor),
          color: draft.fontColor,
          fontFamily: draft.fontFamily,
          fontSize: `${Math.max(10, draft.fontSize)}px`,
          fontStyle: draft.italic ? 'italic' : 'normal',
          fontWeight: draft.bold ? 700 : 400,
        }}
      >
        {officeMessage(messages, 'spreadsheet.formatCells.font.preview')}
      </div>
    </div>
  );
}

function spreadsheetFontPreviewBackground(color: string): string {
  const match = /^#([0-9a-f]{6})$/i.exec(color.trim());
  if (!match?.[1]) return '#f4f6f8';
  const channels = [0, 2, 4].map((offset) =>
    Number.parseInt(match[1].slice(offset, offset + 2), 16),
  );
  const [red = 0, green = 0, blue = 0] = channels.map((channel) => {
    const component = channel / 255;
    return component <= 0.04045
      ? component / 12.92
      : ((component + 0.055) / 1.055) ** 2.4;
  });
  const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
  return luminance >= 0.55 ? '#172033' : '#f4f6f8';
}

function BorderPanel({ source, draft, touched, setDraft, touch }: PanelProps) {
  const messages = useOfficeMessages();
  const hasTarget = (target: SpreadsheetCellBorderFormat['target']) =>
    draft.borders.some((format) => format.target === target);
  const updateBorders = (borders: SpreadsheetCellBorderFormat[]) => {
    touch('borders');
    setDraft((current) => ({ ...current, borders }));
  };
  const setBorderPen = (
    field: 'borderColor' | 'borderStyle',
    value: string,
  ) => {
    setDraft((current) => ({
      ...current,
      [field]: value,
    }));
  };
  return (
    <div className="work-spreadsheet-format-cells-border">
      <div className="work-spreadsheet-format-cells-border-tools">
        <Field label={officeMessage(messages, 'spreadsheet.formatCells.border.style')}>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.border.styleAria')}
            value={draft.borderStyle}
            options={borderStyleOptions(messages)}
            onValueChange={(value) => setBorderPen('borderStyle', value)}
          />
        </Field>
        <Field label={officeMessage(messages, 'spreadsheet.formatCells.border.color')}>
          <OfficeColorPicker
            ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.border.colorAria')}
            value={draft.borderColor}
            onValueChange={(value) => setBorderPen('borderColor', value)}
          />
        </Field>
      </div>
      <div className="work-spreadsheet-format-cells-border-layout">
        <div className="work-spreadsheet-format-cells-border-actions">
          {borderTargets(messages).map(({ target, label }) => (
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
            {officeMessage(messages, 'spreadsheet.formatCells.border.none')}
          </Button>
        </div>
        <BorderPreview draft={draft} />
      </div>
      {source.fields.borders.mixed && !touched.borders && <MixedHint />}
    </div>
  );
}

function BorderPreview({ draft }: { draft: SpreadsheetFormatCellsDraft }) {
  const messages = useOfficeMessages();
  const edge = (
    target: SpreadsheetCellBorderFormat['target'],
    className: string,
  ) => {
    const format = draft.borders.find((candidate) => candidate.target === target);
    if (!format) return null;
    return (
      <span
        className={className}
        data-border-style={format.style}
        style={{ '--cell-border-color': format.color } as CSSProperties}
      />
    );
  };
  return (
    <div
      className="work-spreadsheet-format-cells-border-preview"
      role="img"
      aria-label={officeMessage(messages, 'spreadsheet.formatCells.border.previewAria')}
    >
      {edge('top', 'top')}
      {edge('bottom', 'bottom')}
      {edge('left', 'left')}
      {edge('right', 'right')}
      {edge('diagonalDown', 'diagonal-down')}
      {edge('diagonalUp', 'diagonal-up')}
      <strong>{officeMessage(messages, 'spreadsheet.formatCells.border.previewText')}</strong>
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
  const messages = useOfficeMessages();
  return (
    <div className="work-spreadsheet-format-cells-protection">
      <OfficeCheckbox
        ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.protection.lock')}
        checked={draft.locked}
        indeterminate={source.fields.locked.mixed && !touched.locked}
        onCheckedChange={(locked) => {
          touch('locked');
          setDraft((current) => ({ ...current, locked }));
        }}
      >
        {officeMessage(messages, 'spreadsheet.formatCells.protection.lock')}
      </OfficeCheckbox>
      <OfficeCheckbox
        ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.protection.hide')}
        checked={draft.hidden}
        indeterminate={source.fields.hidden.mixed && !touched.hidden}
        onCheckedChange={(hidden) => {
          touch('hidden');
          setDraft((current) => ({ ...current, hidden }));
        }}
      >
        {officeMessage(messages, 'spreadsheet.formatCells.protection.hide')}
      </OfficeCheckbox>
      <p>{officeMessage(messages, 'spreadsheet.formatCells.protection.note')}</p>
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
  const messages = useOfficeMessages();
  return (
    <p className="work-spreadsheet-format-cells-mixed">
      {officeMessage(messages, 'spreadsheet.formatCells.mixedHint')}
    </p>
  );
}

function withMixedOption<T extends string>(
  options: readonly OfficeSelectOption<T>[],
  mixed: boolean,
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<T | typeof MIXED_VALUE>[] {
  return mixed
    ? [
        {
          value: MIXED_VALUE,
          label: officeMessage(messages, 'spreadsheet.formatCells.mixed'),
          disabled: true,
        },
        ...options,
      ]
    : options;
}

