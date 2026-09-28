import type { CSSProperties } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { documentBorderPresentation } from '../work-document-paragraph-borders';
import type { DocumentRunBorderStyle } from '../work-document-run-border';
import {
  type DocumentFontDialogRunBorderDraft,
  type DocumentFontDialogRunBorderMode,
  type DocumentFontDialogRunBorderSource,
  documentFontDialogRunBorderFromDraft,
} from './document-font-dialog-run-border-model';
import {
  OfficeCheckbox,
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

function runBorderModes(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentFontDialogRunBorderMode;
  label: string;
  disabled?: boolean;
}> {
  return [
    {
      value: 'mixed',
      label: officeMessage(messages, 'document.font.mode.mixed'),
      disabled: true,
    },
    {
      value: 'inherit',
      label: officeMessage(messages, 'document.font.mode.inherit'),
    },
    {
      value: 'none',
      label: officeMessage(messages, 'document.font.mode.noneExplicit'),
    },
    {
      value: 'value',
      label: officeMessage(messages, 'document.font.border.modeValue'),
    },
  ];
}

const RUN_BORDER_STYLE_KEYS = [
  'single',
  'thick',
  'double',
  'dotted',
  'dashed',
  'dotDash',
  'dotDotDash',
  'triple',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
  'wave',
  'doubleWave',
  'dashSmallGap',
  'dashDotStroked',
  'threeDEmboss',
  'threeDEngrave',
  'outset',
  'inset',
] as const satisfies ReadonlyArray<DocumentRunBorderStyle>;

function runBorderStyles(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentRunBorderStyle;
  label: string;
}> {
  return RUN_BORDER_STYLE_KEYS.map((value) => ({
    value,
    label: officeMessage(messages, `document.font.border.style.${value}`),
  }));
}

export function DocumentFontDialogRunBorderSection({
  source,
  draft,
  touched,
  onDraftChange,
  onTouched,
}: {
  source: DocumentFontDialogRunBorderSource;
  draft: DocumentFontDialogRunBorderDraft;
  touched: boolean;
  onDraftChange: (patch: Partial<DocumentFontDialogRunBorderDraft>) => void;
  onTouched: () => void;
}) {
  const messages = useOfficeMessages();
  const enabled = draft.runBorderMode === 'value';
  const update = (patch: Partial<DocumentFontDialogRunBorderDraft>) => {
    onDraftChange(patch);
    onTouched();
  };
  return (
    <fieldset
      className="work-document-font-dialog-run-border"
      aria-label={officeMessage(messages, 'document.font.border.sectionAria')}
    >
      <legend>
        {officeMessage(messages, 'document.font.border.legend')}
      </legend>
      <div className="work-document-font-dialog-field">
        <span>{officeMessage(messages, 'document.font.border.apply')}</span>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'document.font.border.applyAria')}
          value={draft.runBorderMode}
          options={runBorderModes(messages)}
          onValueChange={(runBorderMode) => update({ runBorderMode })}
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>{officeMessage(messages, 'document.font.border.style')}</span>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'document.font.border.styleAria')}
          value={draft.runBorderStyle}
          options={runBorderStyles(messages)}
          disabled={!enabled}
          onValueChange={(runBorderStyle) =>
            update({ runBorderMode: 'value', runBorderStyle })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>{officeMessage(messages, 'document.font.border.color')}</span>
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'document.font.border.colorAria')}
          value={
            draft.runBorderColor === 'auto' ? '#000000' : draft.runBorderColor
          }
          disabled={!enabled}
          resetAction={{
            kind: 'automatic',
            label: officeMessage(messages, 'document.color.automatic'),
            onSelect: () =>
              update({ runBorderMode: 'value', runBorderColor: 'auto' }),
          }}
          onValueChange={(runBorderColor) =>
            update({
              runBorderMode: 'value',
              runBorderColor: runBorderColor as `#${string}`,
            })
          }
        />
      </div>
      <div className="work-document-font-dialog-field">
        <span>{officeMessage(messages, 'document.font.border.width')}</span>
        <span className="work-document-font-dialog-measure">
          <OfficeNumberField
            ariaLabel={officeMessage(messages, 'document.font.border.widthAria')}
            value={draft.runBorderWidthPoints}
            min={0.25}
            max={12}
            step={0.125}
            disabled={!enabled}
            onValueChange={(runBorderWidthPoints) =>
              update({ runBorderMode: 'value', runBorderWidthPoints })
            }
          />
          <span aria-hidden="true">
            {officeMessage(messages, 'document.font.unit.points')}
          </span>
        </span>
      </div>
      <div className="work-document-font-dialog-field">
        <span>{officeMessage(messages, 'document.font.border.spacing')}</span>
        <span className="work-document-font-dialog-measure">
          <OfficeNumberField
            ariaLabel={officeMessage(
              messages,
              'document.font.border.spacingAria',
            )}
            value={draft.runBorderSpacingPoints}
            min={0}
            max={31}
            step={1}
            disabled={!enabled}
            onValueChange={(runBorderSpacingPoints) =>
              update({ runBorderMode: 'value', runBorderSpacingPoints })
            }
          />
          <span aria-hidden="true">
            {officeMessage(messages, 'document.font.unit.points')}
          </span>
        </span>
      </div>
      <div className="work-document-font-dialog-run-border-effects">
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.font.border.shadowAria')}
          checked={draft.runBorderShadow}
          disabled={!enabled}
          onCheckedChange={(runBorderShadow) =>
            update({ runBorderMode: 'value', runBorderShadow })
          }
        >
          {officeMessage(messages, 'document.font.border.shadow')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.font.border.frameAria')}
          checked={draft.runBorderFrame}
          disabled={!enabled}
          onCheckedChange={(runBorderFrame) =>
            update({ runBorderMode: 'value', runBorderFrame })
          }
        >
          {officeMessage(messages, 'document.font.border.frame')}
        </OfficeCheckbox>
      </div>
      {source.mixed && !touched ? (
        <p className="work-document-font-dialog-mixed" role="status">
          {officeMessage(messages, 'document.font.border.mixed')}
        </p>
      ) : null}
    </fieldset>
  );
}

export function documentFontDialogRunBorderPreviewStyle(
  source: DocumentFontDialogRunBorderSource,
  draft: DocumentFontDialogRunBorderDraft,
): CSSProperties {
  const border = documentFontDialogRunBorderFromDraft(source, draft);
  if (!border) return {};
  const presentation = documentBorderPresentation(border);
  return {
    border: `${presentation.width}px ${presentation.style} ${presentation.color}`,
    padding: `${(border.space ?? 0) * (96 / 72)}px`,
    boxDecorationBreak: 'clone',
    WebkitBoxDecorationBreak: 'clone',
    ...(border.shadow && presentation.width > 0
      ? { boxShadow: `2px 2px 0 ${presentation.color}` }
      : {}),
  };
}
