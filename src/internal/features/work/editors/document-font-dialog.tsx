import { type CSSProperties, type FormEvent, useId, useState } from 'react';
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

const characterSpacingModes: ReadonlyArray<{
  value: DocumentCharacterSpacingMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'normal', label: '标准' },
  { value: 'expanded', label: '加宽' },
  { value: 'condensed', label: '紧缩' },
];

const characterPositionModes: ReadonlyArray<{
  value: DocumentCharacterPositionMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'normal', label: '标准' },
  { value: 'raised', label: '提升' },
  { value: 'lowered', label: '降低' },
];

const emphasisMarkModes: ReadonlyArray<{
  value: DocumentEmphasisMarkMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'none', label: '无' },
  { value: 'dot', label: '上方圆点' },
  { value: 'comma', label: '上方逗号' },
  { value: 'circle', label: '上方圆圈' },
  { value: 'underDot', label: '下方圆点' },
];

const openTypeLigatureModes: ReadonlyArray<{
  value: DocumentOpenTypeLigaturesMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'none', label: '无' },
  { value: 'standard', label: '标准' },
  { value: 'contextual', label: '上下文' },
  { value: 'historical', label: '历史' },
  { value: 'discretional', label: '任意' },
  { value: 'standardContextual', label: '标准和上下文' },
  { value: 'standardHistorical', label: '标准和历史' },
  { value: 'contextualHistorical', label: '上下文和历史' },
  { value: 'standardDiscretional', label: '标准和任意' },
  { value: 'contextualDiscretional', label: '上下文和任意' },
  { value: 'historicalDiscretional', label: '历史和任意' },
  { value: 'standardContextualHistorical', label: '标准、上下文和历史' },
  { value: 'standardContextualDiscretional', label: '标准、上下文和任意' },
  { value: 'standardHistoricalDiscretional', label: '标准、历史和任意' },
  { value: 'contextualHistoricalDiscretional', label: '上下文、历史和任意' },
  { value: 'all', label: '全部' },
];

const openTypeNumberFormModes: ReadonlyArray<{
  value: DocumentOpenTypeNumberFormMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'default', label: '默认' },
  { value: 'lining', label: '等高数字' },
  { value: 'oldStyle', label: '旧式数字' },
];

const openTypeNumberSpacingModes: ReadonlyArray<{
  value: DocumentOpenTypeNumberSpacingMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'default', label: '默认' },
  { value: 'proportional', label: '比例宽度' },
  { value: 'tabular', label: '等宽' },
];

const openTypeStylisticSetModes: ReadonlyArray<{
  value: DocumentOpenTypeStylisticSetsMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'multiple', label: '多个样式集（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'none', label: '无' },
  ...Array.from({ length: 20 }, (_, index) => ({
    value: `set-${index + 1}` as const,
    label: `样式集 ${index + 1}`,
  })),
];

const openTypeContextualAlternatesModes: ReadonlyArray<{
  value: DocumentOpenTypeContextualAlternatesMode;
  label: string;
  disabled?: boolean;
}> = [
  { value: 'mixed', label: '混合（保持不变）', disabled: true },
  { value: 'inherit', label: '跟随样式' },
  { value: 'enabled', label: '启用' },
  { value: 'disabled', label: '禁用' },
];

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
  const error = documentFontDialogDraftError(draft);
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

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (!hasChanges || error) return;
    if (onApply(patch)) onClose();
  };

  return (
    <Dialog
      title="字体高级设置"
      description={fontDialogDescription(source)}
      className="work-document-font-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!hasChanges || Boolean(error)}
          >
            应用
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-document-font-dialog-script-fonts">
          <legend>按文字系统设置字体</legend>
          <div className="work-document-font-dialog-field">
            <span>拉丁文字</span>
            <OfficeSelect
              ariaLabel="拉丁文字字体"
              initialFocus
              value={draft.latinFont}
              options={scriptFontOptions(draft.latinFont, layoutFonts)}
              onValueChange={(latinFont) => {
                setDraft((current) => ({ ...current, latinFont }));
                setLatinFontTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field">
            <span>东亚文字</span>
            <OfficeSelect
              ariaLabel="东亚文字字体"
              value={draft.eastAsiaFont}
              options={scriptFontOptions(draft.eastAsiaFont, layoutFonts)}
              onValueChange={(eastAsiaFont) => {
                setDraft((current) => ({ ...current, eastAsiaFont }));
                setEastAsiaFontTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field">
            <span>复杂文字</span>
            <OfficeSelect
              ariaLabel="复杂文字字体"
              value={draft.complexScriptFont}
              options={scriptFontOptions(draft.complexScriptFont, layoutFonts)}
              onValueChange={(complexScriptFont) => {
                setDraft((current) => ({ ...current, complexScriptFont }));
                setComplexScriptFontTouched(true);
              }}
            />
          </div>
          {source.latinFont.mixed && !latinFontTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含不同的拉丁文字字体。选择字体后才会统一修改。
            </p>
          )}
          {source.eastAsiaFont.mixed && !eastAsiaFontTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含不同的东亚文字字体。选择字体后才会统一修改。
            </p>
          )}
          {source.complexScriptFont.mixed && !complexScriptFontTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含不同的复杂文字字体。选择字体后才会统一修改。
            </p>
          )}
        </fieldset>
        <fieldset className="work-document-font-dialog-spacing">
          <legend>字符缩放、间距、字距调整、位置与文字效果</legend>
          <div className="work-document-font-dialog-field">
            <span>缩放</span>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel="字符缩放比例（%）"
                value={draft.characterScalePercent}
                min={1}
                max={600}
                step={1}
                placeholder={
                  draft.characterScaleMode === 'mixed' ? '混合' : undefined
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
            <span>间距</span>
            <OfficeSelect
              ariaLabel="字符间距"
              value={draft.characterSpacingMode}
              options={characterSpacingModes}
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
            <span>间距值</span>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel="间距值（磅）"
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
              <span aria-hidden="true">磅</span>
            </span>
          </div>
          <div className="work-document-font-dialog-field">
            <OfficeCheckbox
              ariaLabel="为字号达到以下值的字体调整字距"
              checked={draft.kerningEnabled}
              indeterminate={source.kerningThreshold.mixed && !kerningTouched}
              onCheckedChange={(kerningEnabled) => {
                setDraft((current) => ({ ...current, kerningEnabled }));
                setKerningTouched(true);
              }}
            >
              为字号达到以下值的字体调整字距
            </OfficeCheckbox>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel="字距调整阈值（磅）"
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
              <span aria-hidden="true">磅</span>
            </span>
          </div>
          <div className="work-document-font-dialog-field">
            <span>位置</span>
            <OfficeSelect
              ariaLabel="字符位置"
              value={draft.characterPositionMode}
              options={characterPositionModes}
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
            <span>位置值</span>
            <span className="work-document-font-dialog-measure">
              <OfficeNumberField
                ariaLabel="位置值（磅）"
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
              <span aria-hidden="true">磅</span>
            </span>
          </div>
          <div className="work-document-font-dialog-field">
            <span>着重号</span>
            <OfficeSelect
              ariaLabel="着重号"
              value={draft.emphasisMark}
              options={emphasisMarkModes}
              onValueChange={(emphasisMark) => {
                setDraft((current) => ({ ...current, emphasisMark }));
                setEmphasisTouched(true);
              }}
            />
          </div>
          <div className="work-document-font-dialog-field work-document-font-dialog-hidden-text">
            <OfficeCheckbox
              ariaLabel="隐藏文字"
              checked={draft.hiddenText}
              indeterminate={source.hiddenText.mixed && !hiddenTextTouched}
              onCheckedChange={(hiddenText) => {
                setDraft((current) => ({ ...current, hiddenText }));
                setHiddenTextTouched(true);
              }}
            >
              隐藏文字
            </OfficeCheckbox>
          </div>
          <fieldset
            className="work-document-font-dialog-legacy-effects"
            aria-label="文字效果"
          >
            <OfficeCheckbox
              ariaLabel="空心"
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
              空心
            </OfficeCheckbox>
            <OfficeCheckbox
              ariaLabel="阴影"
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
              阴影
            </OfficeCheckbox>
            <OfficeCheckbox
              ariaLabel="阳文"
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
              阳文
            </OfficeCheckbox>
            <OfficeCheckbox
              ariaLabel="阴文"
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
              阴文
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
            <legend>OpenType 排版</legend>
            <div className="work-document-font-dialog-field">
              <span>连字</span>
              <OfficeSelect
                ariaLabel="OpenType 连字"
                value={draft.openTypeLigatures}
                options={openTypeLigatureModes}
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
              <span>数字字形</span>
              <OfficeSelect
                ariaLabel="OpenType 数字字形"
                value={draft.openTypeNumberForm}
                options={openTypeNumberFormModes}
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
              <span>数字间距</span>
              <OfficeSelect
                ariaLabel="OpenType 数字间距"
                value={draft.openTypeNumberSpacing}
                options={openTypeNumberSpacingModes}
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
              <span>样式集</span>
              <OfficeSelect
                ariaLabel="OpenType 样式集"
                value={draft.openTypeStylisticSets}
                options={openTypeStylisticSetModes}
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
              <span>上下文替代</span>
              <OfficeSelect
                ariaLabel="OpenType 上下文替代"
                value={draft.openTypeContextualAlternates}
                options={openTypeContextualAlternatesModes}
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
                当前选区包含不同的 OpenType
                排版设置。仅修改所选项目，其他原生设置保持不变。
              </p>
            )}
          </fieldset>
          {source.characterScale.mixed && !characterScaleTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含多种字符缩放比例。输入缩放比例后才会统一修改。
            </p>
          )}
          {source.characterSpacing.mixed && !characterSpacingTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含多种字符间距。选择一种间距后才会统一修改。
            </p>
          )}
          {source.characterPosition.mixed && !characterPositionTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含多种字符位置。选择一种位置后才会统一修改。
            </p>
          )}
          {source.kerningThreshold.mixed && !kerningTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含不同的字距调整设置。勾选或取消后才会统一修改。
            </p>
          )}
          {source.emphasisMark.mixed && !emphasisTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含不同的着重号。选择一种设置后才会统一修改。
            </p>
          )}
          {source.hiddenText.mixed && !hiddenTextTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区同时包含隐藏和可见文字。勾选或取消后才会统一修改。
            </p>
          )}
          {((source.legacyTextOutline.mixed && !legacyTextOutlineTouched) ||
            (source.legacyTextShadow.mixed && !legacyTextShadowTouched) ||
            (source.legacyTextEmboss.mixed && !legacyTextEmbossTouched) ||
            (source.legacyTextImprint.mixed && !legacyTextImprintTouched)) && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含不同的空心、阴影、阳文或阴文设置。勾选或取消对应选项后才会统一修改。
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
          aria-label="字符高级格式预览"
        >
          <span>预览</span>
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
  value: string,
  layoutFonts: readonly WorkDocumentLayoutFont[],
): readonly OfficeSelectOption[] {
  const catalogValue =
    value === 'mixed' || value === 'inherit' ? 'default' : value;
  return [
    { value: 'mixed', label: '混合（保持不变）', disabled: true },
    { value: 'inherit', label: '跟随样式' },
    ...documentFontFamilyOptionsForValue(catalogValue, layoutFonts).filter(
      (option) => option.value !== 'default',
    ),
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

function fontDialogDescription(source: DocumentFontDialogSource): string {
  return source.selectedCharacters
    ? `分别设置当前选中内容的拉丁、东亚和复杂文字字体，以及原生字符缩放、间距、字距调整阈值、位置、文字效果和 OpenType 排版（${source.selectedCharacters} 个字符）。`
    : '分别设置当前位置后续输入文字的拉丁、东亚和复杂文字字体，以及原生字符缩放、间距、字距调整阈值、位置、文字效果和 OpenType 排版。';
}
