import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
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
  type DocumentFontDialogPatch,
  type DocumentFontDialogSource,
} from './document-font-dialog-model';

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

export function DocumentFontDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  source: DocumentFontDialogSource;
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
  const formId = useId();
  const error = documentFontDialogDraftError(draft);
  const patch = documentFontDialogPatch(
    source,
    draft,
    characterScaleTouched,
    characterSpacingTouched,
    characterPositionTouched,
    kerningTouched,
  );
  const hasChanges = Object.keys(patch).length > 0;
  const previewScale = previewCharacterScale(draft);
  const previewSpacing = previewCharacterSpacing(draft);
  const previewPosition = previewCharacterPosition(draft);
  const previewKerning = previewDocumentKerning(draft, source.fontSize);

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
        <fieldset className="work-document-font-dialog-spacing">
          <legend>字符缩放、间距、字距调整与位置</legend>
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
              initialFocus
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
          {(characterScaleTouched ||
            characterSpacingTouched ||
            characterPositionTouched ||
            kerningTouched) &&
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
            }}
          >
            <span style={{ verticalAlign: `${previewPosition}pt` }}>
              {source.previewText}
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

function fontDialogDescription(source: DocumentFontDialogSource): string {
  return source.selectedCharacters
    ? `精确设置当前选中内容的原生字符缩放、间距、字距调整阈值和位置（${source.selectedCharacters} 个字符）。`
    : '设置当前位置后续输入文字的原生字符缩放、间距、字距调整阈值和位置。';
}
