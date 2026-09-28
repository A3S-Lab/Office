import { type CSSProperties, type FormEvent, useId, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { Button, Dialog } from '../../../design-system/primitives';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import { documentScriptFontSegments } from '../work-document-script-fonts';
import { documentLegacyTextEffectsCss } from '../work-document-legacy-text-effects';
import { documentOpenTypeCssProperties } from '../work-document-opentype';
import {
  documentKerningIsEffective,
  DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS,
} from '../work-document-kerning';
import {
  OfficeCheckbox,
  OfficeNumberField,
  OfficeSelect,
} from './office-controls';
import {
  createDocumentFontDialogDraft,
  documentFontDialogDraftError,
  documentFontDialogPatch,
  type DocumentCharacterPositionMode,
  type DocumentCharacterSpacingMode,
  type DocumentEmphasisMarkMode,
  type DocumentFontDialogPatch,
  type DocumentFontDialogSource,
} from './document-font-dialog-model';
import { documentFontFamilyOptionsForValue } from './document-formatting-options';
import type { OfficeSelectOption } from './office-select';
import {
  DocumentFontDialogRunBorderSection,
  documentFontDialogRunBorderPreviewStyle,
} from './document-font-dialog-run-border-section';
import { DocumentFontDialogRunShadingSection } from './document-font-dialog-run-shading-section';
import { documentFontDialogRunShadingPreviewStyle } from './document-font-dialog-run-shading-model';
import {
  documentFontDialogOpenTypePreviewFeatures,
  type DocumentOpenTypeContextualAlternatesMode,
  type DocumentOpenTypeLigaturesMode,
  type DocumentOpenTypeNumberFormMode,
  type DocumentOpenTypeNumberSpacingMode,
  type DocumentOpenTypeStylisticSetsMode,
} from './document-font-dialog-opentype-model';
import { useOfficeMessages } from './office-messages-context';

function characterSpacingModes(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentCharacterSpacingMode;
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
      value: 'normal',
      label: officeMessage(messages, 'document.font.mode.normal'),
    },
    {
      value: 'expanded',
      label: officeMessage(messages, 'document.font.spacingMode.expanded'),
    },
    {
      value: 'condensed',
      label: officeMessage(messages, 'document.font.spacingMode.condensed'),
    },
  ];
}

function characterPositionModes(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentCharacterPositionMode;
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
      value: 'normal',
      label: officeMessage(messages, 'document.font.mode.normal'),
    },
    {
      value: 'raised',
      label: officeMessage(messages, 'document.font.positionMode.raised'),
    },
    {
      value: 'lowered',
      label: officeMessage(messages, 'document.font.positionMode.lowered'),
    },
  ];
}

function emphasisMarkModes(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentEmphasisMarkMode;
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
      label: officeMessage(messages, 'document.font.mode.none'),
    },
    {
      value: 'dot',
      label: officeMessage(messages, 'document.font.emphasis.dot'),
    },
    {
      value: 'comma',
      label: officeMessage(messages, 'document.font.emphasis.comma'),
    },
    {
      value: 'circle',
      label: officeMessage(messages, 'document.font.emphasis.circle'),
    },
    {
      value: 'underDot',
      label: officeMessage(messages, 'document.font.emphasis.underDot'),
    },
  ];
}

function openTypeLigatureModes(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentOpenTypeLigaturesMode;
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
      label: officeMessage(messages, 'document.font.mode.none'),
    },
    {
      value: 'standard',
      label: officeMessage(messages, 'document.font.mode.normal'),
    },
    {
      value: 'contextual',
      label: officeMessage(messages, 'document.font.ligatures.contextual'),
    },
    {
      value: 'historical',
      label: officeMessage(messages, 'document.font.ligatures.historical'),
    },
    {
      value: 'discretional',
      label: officeMessage(messages, 'document.font.ligatures.discretional'),
    },
    {
      value: 'standardContextual',
      label: officeMessage(messages, 'document.font.ligatures.standardContextual'),
    },
    {
      value: 'standardHistorical',
      label: officeMessage(messages, 'document.font.ligatures.standardHistorical'),
    },
    {
      value: 'contextualHistorical',
      label: officeMessage(
        messages,
        'document.font.ligatures.contextualHistorical',
      ),
    },
    {
      value: 'standardDiscretional',
      label: officeMessage(
        messages,
        'document.font.ligatures.standardDiscretional',
      ),
    },
    {
      value: 'contextualDiscretional',
      label: officeMessage(
        messages,
        'document.font.ligatures.contextualDiscretional',
      ),
    },
    {
      value: 'historicalDiscretional',
      label: officeMessage(
        messages,
        'document.font.ligatures.historicalDiscretional',
      ),
    },
    {
      value: 'standardContextualHistorical',
      label: officeMessage(
        messages,
        'document.font.ligatures.standardContextualHistorical',
      ),
    },
    {
      value: 'standardContextualDiscretional',
      label: officeMessage(
        messages,
        'document.font.ligatures.standardContextualDiscretional',
      ),
    },
    {
      value: 'standardHistoricalDiscretional',
      label: officeMessage(
        messages,
        'document.font.ligatures.standardHistoricalDiscretional',
      ),
    },
    {
      value: 'contextualHistoricalDiscretional',
      label: officeMessage(
        messages,
        'document.font.ligatures.contextualHistoricalDiscretional',
      ),
    },
    {
      value: 'all',
      label: officeMessage(messages, 'document.font.ligatures.all'),
    },
  ];
}

function openTypeNumberFormModes(messages: OfficeMessageCatalog): ReadonlyArray<{
  value: DocumentOpenTypeNumberFormMode;
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
      value: 'default',
      label: officeMessage(messages, 'document.font.mode.default'),
    },
    {
      value: 'lining',
      label: officeMessage(messages, 'document.font.numberForm.lining'),
    },
    {
      value: 'oldStyle',
      label: officeMessage(messages, 'document.font.numberForm.oldStyle'),
    },
  ];
}

function openTypeNumberSpacingModes(
  messages: OfficeMessageCatalog,
): ReadonlyArray<{
  value: DocumentOpenTypeNumberSpacingMode;
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
      value: 'default',
      label: officeMessage(messages, 'document.font.mode.default'),
    },
    {
      value: 'proportional',
      label: officeMessage(messages, 'document.font.numberSpacing.proportional'),
    },
    {
      value: 'tabular',
      label: officeMessage(messages, 'document.font.numberSpacing.tabular'),
    },
  ];
}

function openTypeStylisticSetModes(
  messages: OfficeMessageCatalog,
): ReadonlyArray<{
  value: DocumentOpenTypeStylisticSetsMode;
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
      value: 'multiple',
      label: officeMessage(messages, 'document.font.stylisticSets.multiple'),
      disabled: true,
    },
    {
      value: 'inherit',
      label: officeMessage(messages, 'document.font.mode.inherit'),
    },
    {
      value: 'none',
      label: officeMessage(messages, 'document.font.mode.none'),
    },
    ...Array.from({ length: 20 }, (_, index) => ({
      value: `set-${index + 1}` as const,
      label: officeMessage(messages, 'document.font.openType.stylisticSetN', {
        n: String(index + 1),
      }),
    })),
  ];
}

function openTypeContextualAlternatesModes(
  messages: OfficeMessageCatalog,
): ReadonlyArray<{
  value: DocumentOpenTypeContextualAlternatesMode;
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
      value: 'enabled',
      label: officeMessage(messages, 'document.font.mode.enabled'),
    },
    {
      value: 'disabled',
      label: officeMessage(messages, 'document.font.mode.disabled'),
    },
  ];
}

export function DocumentFontDialog({
  source,
  layoutFonts = [],
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  source: DocumentFontDialogSource;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (patch: DocumentFontDialogPatch) => boolean;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const spacingModes = characterSpacingModes(messages);
  const positionModes = characterPositionModes(messages);
  const emphasisModes = emphasisMarkModes(messages);
  const ligatureModes = openTypeLigatureModes(messages);
  const numberFormModes = openTypeNumberFormModes(messages);
  const numberSpacingModes = openTypeNumberSpacingModes(messages);
  const stylisticSetModes = openTypeStylisticSetModes(messages);
  const contextualAlternateModes = openTypeContextualAlternatesModes(messages);
  const [draft, setDraft] = useState(() =>
    createDocumentFontDialogDraft(source),
  );
  const [characterScaleTouched, setCharacterScaleTouched] = useState(false);
  const [characterSpacingTouched, setCharacterSpacingTouched] = useState(false);
  const [characterPositionTouched, setCharacterPositionTouched] =
    useState(false);
  const [kerningTouched, setKerningTouched] = useState(false);
  const [emphasisTouched, setEmphasisTouched] = useState(false);
  const [hiddenTextTouched, setHiddenTextTouched] = useState(false);
  const [legacyTextOutlineTouched, setLegacyTextOutlineTouched] =
    useState(false);
  const [legacyTextShadowTouched, setLegacyTextShadowTouched] = useState(false);
  const [legacyTextEmbossTouched, setLegacyTextEmbossTouched] = useState(false);
  const [legacyTextImprintTouched, setLegacyTextImprintTouched] =
    useState(false);
  const [runBorderTouched, setRunBorderTouched] = useState(false);
  const [runShadingTouched, setRunShadingTouched] = useState(false);
  const [latinFontTouched, setLatinFontTouched] = useState(false);
  const [eastAsiaFontTouched, setEastAsiaFontTouched] = useState(false);
  const [complexScriptFontTouched, setComplexScriptFontTouched] =
    useState(false);
  const [openTypeLigaturesTouched, setOpenTypeLigaturesTouched] =
    useState(false);
  const [openTypeNumberFormTouched, setOpenTypeNumberFormTouched] =
    useState(false);
  const [openTypeNumberSpacingTouched, setOpenTypeNumberSpacingTouched] =
    useState(false);
  const [openTypeStylisticSetsTouched, setOpenTypeStylisticSetsTouched] =
    useState(false);
  const [
    openTypeContextualAlternatesTouched,
    setOpenTypeContextualAlternatesTouched,
  ] = useState(false);
  const formId = useId();
  const error = documentFontDialogDraftError(draft, messages);
  const patch = documentFontDialogPatch(source, draft, {
    characterPosition: characterPositionTouched,
    characterScale: characterScaleTouched,
    characterSpacing: characterSpacingTouched,
    complexScriptFont: complexScriptFontTouched,
    eastAsiaFont: eastAsiaFontTouched,
    emphasisMark: emphasisTouched,
    hiddenText: hiddenTextTouched,
    legacyTextOutline: legacyTextOutlineTouched,
    legacyTextShadow: legacyTextShadowTouched,
    legacyTextEmboss: legacyTextEmbossTouched,
    legacyTextImprint: legacyTextImprintTouched,
    runBorder: runBorderTouched,
    runShading: runShadingTouched,
    kerning: kerningTouched,
    latinFont: latinFontTouched,
    openTypeLigatures: openTypeLigaturesTouched,
    openTypeNumberForm: openTypeNumberFormTouched,
    openTypeNumberSpacing: openTypeNumberSpacingTouched,
    openTypeStylisticSets: openTypeStylisticSetsTouched,
    openTypeContextualAlternates: openTypeContextualAlternatesTouched,
  });
  const hasChanges = Object.keys(patch).length > 0;
  const previewScale = previewCharacterScale(draft);
  const previewSpacing = previewCharacterSpacing(draft);
  const previewPosition = previewCharacterPosition(draft);
  const previewKerning = previewDocumentKerning(draft, source.fontSize);
  const previewEmphasis = previewDocumentEmphasis(draft.emphasisMark);
  const previewLegacyTextEffects = documentLegacyTextEffectsCss({
    outline: draft.legacyTextOutline,
    shadow: draft.legacyTextShadow,
    emboss: draft.legacyTextEmboss,
    imprint: draft.legacyTextImprint,
  });
  const previewRunBorder = documentFontDialogRunBorderPreviewStyle(
    source.runBorder,
    draft,
  );
  const previewRunShading = documentFontDialogRunShadingPreviewStyle(
    source.runShading,
    draft,
  );
  const previewOpenType = documentOpenTypeCssProperties(
    documentFontDialogOpenTypePreviewFeatures(source, draft),
  );

  const resetDraft = () => {
    setDraft(createDocumentFontDialogDraft(source));
    setCharacterScaleTouched(false);
    setCharacterSpacingTouched(false);
    setCharacterPositionTouched(false);
    setKerningTouched(false);
    setEmphasisTouched(false);
    setHiddenTextTouched(false);
    setLegacyTextOutlineTouched(false);
    setLegacyTextShadowTouched(false);
    setLegacyTextEmbossTouched(false);
    setLegacyTextImprintTouched(false);
    setRunBorderTouched(false);
    setRunShadingTouched(false);
    setLatinFontTouched(false);
    setEastAsiaFontTouched(false);
    setComplexScriptFontTouched(false);
    setOpenTypeLigaturesTouched(false);
    setOpenTypeNumberFormTouched(false);
    setOpenTypeNumberSpacingTouched(false);
    setOpenTypeStylisticSetsTouched(false);
    setOpenTypeContextualAlternatesTouched(false);
  };

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (!hasChanges || error) return;
    if (onApply(patch)) onClose();
  };

  return (
    <Dialog
      title={officeMessage(messages, 'document.font.dialog.title')}
      description={fontDialogDescription(messages, source)}
      className="work-document-font-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      onEscape={() => {
        if (hasChanges) {
          resetDraft();
          return;
        }
        onClose();
      }}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'document.font.dialog.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!hasChanges || Boolean(error)}
          >
            {officeMessage(messages, 'document.font.dialog.apply')}
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-document-font-dialog-script-fonts">
          <legend>{officeMessage(messages, 'document.font.script.legend')}</legend>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.script.latin')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.font.script.latinAria')}
              initialFocus
              value={draft.latinFont}
              options={scriptFontOptions(messages, draft.latinFont, layoutFonts)}
              onValueChange={(latinFont) => {
                setDraft((current) => ({ ...current, latinFont }));
                setLatinFontTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.script.eastAsia')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.font.script.eastAsiaAria')}
              value={draft.eastAsiaFont}
              options={scriptFontOptions(messages, draft.eastAsiaFont, layoutFonts)}
              onValueChange={(eastAsiaFont) => {
                setDraft((current) => ({ ...current, eastAsiaFont }));
                setEastAsiaFontTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.script.complex')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.font.script.complexAria')}
              value={draft.complexScriptFont}
              options={scriptFontOptions(messages, draft.complexScriptFont, layoutFonts)}
              onValueChange={(complexScriptFont) => {
                setDraft((current) => ({ ...current, complexScriptFont }));
                setComplexScriptFontTouched(true);
              }}
            />
          </div>
          {source.latinFont.mixed && !latinFontTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.script.mixedLatin')}
            </p>
          )}
          {source.eastAsiaFont.mixed && !eastAsiaFontTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.script.mixedEastAsia')}
            </p>
          )}
          {source.complexScriptFont.mixed && !complexScriptFontTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.script.mixedComplex')}
            </p>
          )}
        </fieldset>
        <fieldset className="work-document-font-dialog-spacing">
          <legend>{officeMessage(messages, 'document.font.spacing.legend')}</legend>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.spacing.scale')}</span>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel={officeMessage(messages, 'document.font.spacing.scaleAria')}
                value={draft.characterScalePercent}
                min={1}
                max={600}
                step={1}
                placeholder={
                  draft.characterScaleMode === 'mixed'
                    ? officeMessage(messages, 'document.font.spacing.mixedPlaceholder')
                    : undefined
                }
                validationInvalid={Boolean(error)}
                onValueChange={(characterScalePercent) => {
                  setDraft((current) => ({
                    ...current,
                    characterScaleMode: 'value',
                    characterScalePercent,
                  }));
                  setCharacterScaleTouched(true);
                }}
              />
              <span aria-hidden="true">%</span>
            </span>
          </div>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.spacing.spacing')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.font.spacing.spacingAria')}
              value={draft.characterSpacingMode}
              options={spacingModes}
              onValueChange={(characterSpacingMode) => {
                setDraft((current) => ({
                  ...current,
                  characterSpacingMode,
                }));
                setCharacterSpacingTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.spacing.spacingValue')}</span>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel={officeMessage(messages, 'document.font.spacing.spacingValueAria')}
                value={draft.characterSpacingPoints}
                min={0.05}
                max={1584}
                step={0.05}
                disabled={
                  draft.characterSpacingMode === 'mixed' ||
                  draft.characterSpacingMode === 'normal'
                }
                validationInvalid={Boolean(error)}
                onValueChange={(characterSpacingPoints) => {
                  setDraft((current) => ({
                    ...current,
                    characterSpacingPoints,
                  }));
                  setCharacterSpacingTouched(true);
                }}
              />
              <span aria-hidden="true">{officeMessage(messages, 'document.font.unit.points')}</span>
            </span>
          </div>
          <div className="work-document-font-dialog-field">
            <OfficeCheckbox
              ariaLabel={officeMessage(messages, 'document.font.kerning.label')}
              checked={draft.kerningEnabled}
              indeterminate={source.kerningThreshold.mixed && !kerningTouched}
              onCheckedChange={(kerningEnabled) => {
                setDraft((current) => ({ ...current, kerningEnabled }));
                setKerningTouched(true);
              }}
            >
              {officeMessage(messages, 'document.font.kerning.label')}
            </OfficeCheckbox>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel={officeMessage(messages, 'document.font.kerning.thresholdAria')}
                value={draft.kerningThresholdPoints}
                min={0}
                max={DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS / 2}
                step={0.5}
                disabled={!draft.kerningEnabled}
                validationInvalid={Boolean(error)}
                onValueChange={(kerningThresholdPoints) => {
                  setDraft((current) => ({
                    ...current,
                    kerningEnabled: true,
                    kerningThresholdPoints,
                  }));
                  setKerningTouched(true);
                }}
              />
              <span aria-hidden="true">{officeMessage(messages, 'document.font.unit.points')}</span>
            </span>
          </div>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.position')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.font.positionAria')}
              value={draft.characterPositionMode}
              options={positionModes}
              onValueChange={(characterPositionMode) => {
                setDraft((current) => ({
                  ...current,
                  characterPositionMode,
                }));
                setCharacterPositionTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.positionValue')}</span>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel={officeMessage(messages, 'document.font.positionValueAria')}
                value={draft.characterPositionPoints}
                min={0.5}
                max={1584}
                step={0.5}
                disabled={
                  draft.characterPositionMode === 'mixed' ||
                  draft.characterPositionMode === 'normal'
                }
                validationInvalid={Boolean(error)}
                onValueChange={(characterPositionPoints) => {
                  setDraft((current) => ({
                    ...current,
                    characterPositionPoints,
                  }));
                  setCharacterPositionTouched(true);
                }}
              />
              <span aria-hidden="true">{officeMessage(messages, 'document.font.unit.points')}</span>
            </span>
          </div>
          <div className="work-document-font-dialog-field">
            <span>{officeMessage(messages, 'document.font.emphasis')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.font.emphasisAria')}
              value={draft.emphasisMark}
              options={emphasisModes}
              onValueChange={(emphasisMark) => {
                setDraft((current) => ({ ...current, emphasisMark }));
                setEmphasisTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field work-document-font-dialog-hidden-text">
            <OfficeCheckbox
              ariaLabel={officeMessage(messages, 'document.font.hiddenText')}
              checked={draft.hiddenText}
              indeterminate={source.hiddenText.mixed && !hiddenTextTouched}
              onCheckedChange={(hiddenText) => {
                setDraft((current) => ({ ...current, hiddenText }));
                setHiddenTextTouched(true);
              }}
            >
              {officeMessage(messages, 'document.font.hiddenText')}
            </OfficeCheckbox>
          </div>
          <fieldset
            className="work-document-font-dialog-legacy-effects"
            aria-label={officeMessage(messages, 'document.font.effects.groupAria')}
          >
            <OfficeCheckbox
              ariaLabel={officeMessage(messages, 'document.font.effects.outline')}
              checked={draft.legacyTextOutline}
              indeterminate={
                source.legacyTextOutline.mixed && !legacyTextOutlineTouched
              }
              onCheckedChange={(legacyTextOutline) => {
                setDraft((current) =>
                  legacyTextOutline
                    ? {
                        ...current,
                        legacyTextOutline,
                        legacyTextEmboss: false,
                        legacyTextImprint: false,
                      }
                    : { ...current, legacyTextOutline },
                );
                setLegacyTextOutlineTouched(true);
                if (legacyTextOutline) {
                  setLegacyTextEmbossTouched(true);
                  setLegacyTextImprintTouched(true);
                }
              }}
            >
              {officeMessage(messages, 'document.font.effects.outline')}
            </OfficeCheckbox>
            <OfficeCheckbox
              ariaLabel={officeMessage(messages, 'document.font.effects.shadow')}
              checked={draft.legacyTextShadow}
              indeterminate={
                source.legacyTextShadow.mixed && !legacyTextShadowTouched
              }
              onCheckedChange={(legacyTextShadow) => {
                setDraft((current) =>
                  legacyTextShadow
                    ? {
                        ...current,
                        legacyTextShadow,
                        legacyTextEmboss: false,
                        legacyTextImprint: false,
                      }
                    : { ...current, legacyTextShadow },
                );
                setLegacyTextShadowTouched(true);
                if (legacyTextShadow) {
                  setLegacyTextEmbossTouched(true);
                  setLegacyTextImprintTouched(true);
                }
              }}
            >
              {officeMessage(messages, 'document.font.effects.shadow')}
            </OfficeCheckbox>
            <OfficeCheckbox
              ariaLabel={officeMessage(messages, 'document.font.effects.emboss')}
              checked={draft.legacyTextEmboss}
              indeterminate={
                source.legacyTextEmboss.mixed && !legacyTextEmbossTouched
              }
              onCheckedChange={(legacyTextEmboss) => {
                setDraft((current) =>
                  legacyTextEmboss
                    ? {
                        ...current,
                        legacyTextOutline: false,
                        legacyTextShadow: false,
                        legacyTextEmboss,
                        legacyTextImprint: false,
                      }
                    : { ...current, legacyTextEmboss },
                );
                setLegacyTextEmbossTouched(true);
                if (legacyTextEmboss) {
                  setLegacyTextOutlineTouched(true);
                  setLegacyTextShadowTouched(true);
                  setLegacyTextImprintTouched(true);
                }
              }}
            >
              {officeMessage(messages, 'document.font.effects.emboss')}
            </OfficeCheckbox>
            <OfficeCheckbox
              ariaLabel={officeMessage(messages, 'document.font.effects.imprint')}
              checked={draft.legacyTextImprint}
              indeterminate={
                source.legacyTextImprint.mixed && !legacyTextImprintTouched
              }
              onCheckedChange={(legacyTextImprint) => {
                setDraft((current) =>
                  legacyTextImprint
                    ? {
                        ...current,
                        legacyTextOutline: false,
                        legacyTextShadow: false,
                        legacyTextEmboss: false,
                        legacyTextImprint,
                      }
                    : { ...current, legacyTextImprint },
                );
                setLegacyTextImprintTouched(true);
                if (legacyTextImprint) {
                  setLegacyTextOutlineTouched(true);
                  setLegacyTextShadowTouched(true);
                  setLegacyTextEmbossTouched(true);
                }
              }}
            >
              {officeMessage(messages, 'document.font.effects.imprint')}
            </OfficeCheckbox>
          </fieldset>
          <DocumentFontDialogRunBorderSection
            source={source.runBorder}
            draft={draft}
            touched={runBorderTouched}
            onDraftChange={(runBorderDraft) =>
              setDraft((current) => ({ ...current, ...runBorderDraft }))
            }
            onTouched={() => setRunBorderTouched(true)}
          />
          <DocumentFontDialogRunShadingSection
            source={source.runShading}
            draft={draft}
            touched={runShadingTouched}
            onDraftChange={(runShadingDraft) =>
              setDraft((current) => ({ ...current, ...runShadingDraft }))
            }
            onTouched={() => setRunShadingTouched(true)}
          />
          <fieldset className="work-document-font-dialog-opentype">
            <legend>{officeMessage(messages, 'document.font.openType.legend')}</legend>
            <div className="work-document-font-dialog-field">
              <span>{officeMessage(messages, 'document.font.openType.ligatures')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'document.font.openType.ligaturesAria')}
                value={draft.openTypeLigatures}
                options={ligatureModes}
                onValueChange={(openTypeLigatures) => {
                  setDraft((current) => ({
                    ...current,
                    openTypeLigatures,
                  }));
                  setOpenTypeLigaturesTouched(true);
                }}
              />
            </div>
            <div className="work-document-font-dialog-field">
              <span>{officeMessage(messages, 'document.font.openType.numberForm')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'document.font.openType.numberFormAria')}
                value={draft.openTypeNumberForm}
                options={numberFormModes}
                onValueChange={(openTypeNumberForm) => {
                  setDraft((current) => ({
                    ...current,
                    openTypeNumberForm,
                  }));
                  setOpenTypeNumberFormTouched(true);
                }}
              />
            </div>
            <div className="work-document-font-dialog-field">
              <span>{officeMessage(messages, 'document.font.openType.numberSpacing')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'document.font.openType.numberSpacingAria')}
                value={draft.openTypeNumberSpacing}
                options={numberSpacingModes}
                onValueChange={(openTypeNumberSpacing) => {
                  setDraft((current) => ({
                    ...current,
                    openTypeNumberSpacing,
                  }));
                  setOpenTypeNumberSpacingTouched(true);
                }}
              />
            </div>
            <div className="work-document-font-dialog-field">
              <span>{officeMessage(messages, 'document.font.openType.stylisticSets')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'document.font.openType.stylisticSetsAria')}
                value={draft.openTypeStylisticSets}
                options={stylisticSetModes}
                onValueChange={(openTypeStylisticSets) => {
                  setDraft((current) => ({
                    ...current,
                    openTypeStylisticSets,
                  }));
                  setOpenTypeStylisticSetsTouched(true);
                }}
              />
            </div>
            <div className="work-document-font-dialog-field">
              <span>{officeMessage(messages, 'document.font.openType.contextualAlternates')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'document.font.openType.contextualAlternatesAria')}
                value={draft.openTypeContextualAlternates}
                options={contextualAlternateModes}
                onValueChange={(openTypeContextualAlternates) => {
                  setDraft((current) => ({
                    ...current,
                    openTypeContextualAlternates,
                  }));
                  setOpenTypeContextualAlternatesTouched(true);
                }}
              />
            </div>
            {((source.openTypeLigatures.mixed && !openTypeLigaturesTouched) ||
              (source.openTypeNumberForm.mixed && !openTypeNumberFormTouched) ||
              (source.openTypeNumberSpacing.mixed &&
                !openTypeNumberSpacingTouched) ||
              (source.openTypeStylisticSets.mixed &&
                !openTypeStylisticSetsTouched) ||
              (source.openTypeContextualAlternates.mixed &&
                !openTypeContextualAlternatesTouched)) && (
              <p className="work-document-font-dialog-mixed" role="status">
                {officeMessage(messages, 'document.font.openType.mixed')}
              </p>
            )}
          </fieldset>
          {source.characterScale.mixed && !characterScaleTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.mixed.scale')}
            </p>
          )}
          {source.characterSpacing.mixed && !characterSpacingTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.mixed.spacing')}
            </p>
          )}
          {source.characterPosition.mixed && !characterPositionTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.mixed.position')}
            </p>
          )}
          {source.kerningThreshold.mixed && !kerningTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.mixed.kerning')}
            </p>
          )}
          {source.emphasisMark.mixed && !emphasisTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.mixed.emphasis')}
            </p>
          )}
          {source.hiddenText.mixed && !hiddenTextTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.mixed.hiddenText')}
            </p>
          )}
          {((source.legacyTextOutline.mixed && !legacyTextOutlineTouched) ||
            (source.legacyTextShadow.mixed && !legacyTextShadowTouched) ||
            (source.legacyTextEmboss.mixed && !legacyTextEmbossTouched) ||
            (source.legacyTextImprint.mixed && !legacyTextImprintTouched)) && (
            <p className="work-document-font-dialog-mixed" role="status">
              {officeMessage(messages, 'document.font.mixed.legacyEffects')}
            </p>
          )}
          {(characterScaleTouched ||
            characterSpacingTouched ||
            characterPositionTouched ||
            kerningTouched ||
            emphasisTouched ||
            hiddenTextTouched ||
            legacyTextOutlineTouched ||
            legacyTextShadowTouched ||
            legacyTextEmbossTouched ||
            legacyTextImprintTouched ||
            runBorderTouched ||
            runShadingTouched ||
            openTypeLigaturesTouched ||
            openTypeNumberFormTouched ||
            openTypeNumberSpacingTouched ||
            openTypeStylisticSetsTouched ||
            openTypeContextualAlternatesTouched) &&
            error && (
              <p className="work-document-font-dialog-error" role="alert">
                {error}
              </p>
            )}
        </fieldset>
        <section
          className="work-document-font-dialog-preview"
          aria-label={officeMessage(messages, 'document.font.preview.aria')}
        >
<span>{officeMessage(messages, 'document.font.preview.label')}</span>
          <output
            style={{
              fontFamily: source.fontFamily ?? undefined,
              fontSize: source.fontSize ?? undefined,
              fontStretch: `${previewScale}%`,
              fontKerning: previewKerning,
              letterSpacing: `${previewSpacing}pt`,
              ...previewOpenType,
            }}
          >
            <span
              style={{
                verticalAlign: `${previewPosition}pt`,
                ...previewEmphasis,
                ...previewLegacyTextEffects,
                ...previewRunShading,
                ...previewRunBorder,
                ...(draft.hiddenText
                  ? {
                      textDecorationColor: 'currentColor',
                      textDecorationLine: 'underline',
                      textDecorationStyle: 'dotted',
                      textUnderlineOffset: '0.18em',
                    }
                  : {}),
              }}
            >
              {documentScriptFontSegments(source.previewText).map(
                ({ from, to, slot }) => (
                  <span
                    key={`${from}-${slot}`}
                    style={{
                      fontFamily: previewScriptFontFamily(
                        slot === 'eastAsia'
                          ? draft.eastAsiaFont
                          : slot === 'complexScript'
                            ? draft.complexScriptFont
                            : draft.latinFont,
                        source.fontFamily,
                      ),
                    }}
                  >
                    {source.previewText.slice(from, to)}
                  </span>
                ),
              )}
            </span>
          </output>
        </section>
      </form>
    </Dialog>
  );
}

function previewCharacterScale(
  draft: ReturnType<typeof createDocumentFontDialogDraft>,
): number {
  if (draft.characterScaleMode === 'mixed') return 100;
  const percent = Number(draft.characterScalePercent);
  return Number.isFinite(percent) ? percent : 100;
}

function previewCharacterPosition(
  draft: ReturnType<typeof createDocumentFontDialogDraft>,
): number {
  if (
    draft.characterPositionMode === 'mixed' ||
    draft.characterPositionMode === 'normal'
  ) {
    return 0;
  }
  const points = Number(draft.characterPositionPoints);
  if (!Number.isFinite(points)) return 0;
  return draft.characterPositionMode === 'lowered' ? -points : points;
}

function previewCharacterSpacing(
  draft: ReturnType<typeof createDocumentFontDialogDraft>,
): number {
  if (
    draft.characterSpacingMode === 'mixed' ||
    draft.characterSpacingMode === 'normal'
  ) {
    return 0;
  }
  const points = Number(draft.characterSpacingPoints);
  if (!Number.isFinite(points)) return 0;
  return draft.characterSpacingMode === 'condensed' ? -points : points;
}

function previewDocumentKerning(
  draft: ReturnType<typeof createDocumentFontDialogDraft>,
  fontSize: string | null,
): 'none' | 'normal' {
  if (!draft.kerningEnabled) return 'none';
  const points = Number(draft.kerningThresholdPoints);
  if (!Number.isFinite(points)) return 'none';
  return documentKerningIsEffective(points * 2, fontSize) ? 'normal' : 'none';
}

function previewDocumentEmphasis(
  emphasisMark: DocumentEmphasisMarkMode,
): Pick<
  CSSProperties,
  | 'textEmphasisPosition'
  | 'textEmphasisStyle'
  | 'WebkitTextEmphasisPosition'
  | 'WebkitTextEmphasisStyle'
> {
  if (
    emphasisMark === 'inherit' ||
    emphasisMark === 'mixed' ||
    emphasisMark === 'none'
  ) {
    return {
      textEmphasisStyle: 'none',
      WebkitTextEmphasisStyle: 'none',
    };
  }
  const style =
    emphasisMark === 'comma'
      ? '","'
      : emphasisMark === 'circle'
        ? 'open circle'
        : 'filled dot';
  const position = emphasisMark === 'underDot' ? 'under right' : 'over right';
  return {
    textEmphasisStyle: style,
    textEmphasisPosition: position,
    WebkitTextEmphasisStyle: style,
    WebkitTextEmphasisPosition: position,
  };
}

function scriptFontOptions(
  messages: OfficeMessageCatalog,
  value: string,
  layoutFonts: readonly WorkDocumentLayoutFont[],
): readonly OfficeSelectOption[] {
  const catalogValue =
    value === 'mixed' || value === 'inherit' ? 'default' : value;
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
    ...documentFontFamilyOptionsForValue(
      catalogValue,
      layoutFonts,
      messages,
    ).filter((option) => option.value !== 'default'),
  ];
}

function previewScriptFontFamily(
  value: string,
  fallback: string | null,
): string | undefined {
  return value === 'mixed' || value === 'inherit'
    ? (fallback ?? undefined)
    : value;
}

function fontDialogDescription(
  messages: OfficeMessageCatalog,
  source: DocumentFontDialogSource,
): string {
  return source.selectedCharacters
    ? officeMessage(messages, 'document.font.dialog.description.selection', {
        count: String(source.selectedCharacters),
      })
    : officeMessage(messages, 'document.font.dialog.description.caret');
}
