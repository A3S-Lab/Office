import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeNumberField, OfficeSelect } from './office-controls';
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
  const [characterSpacingTouched, setCharacterSpacingTouched] = useState(false);
  const [characterPositionTouched, setCharacterPositionTouched] =
    useState(false);
  const formId = useId();
  const error = documentFontDialogDraftError(draft);
  const patch = documentFontDialogPatch(
    source,
    draft,
    characterSpacingTouched,
    characterPositionTouched,
  );
  const hasChanges = Object.keys(patch).length > 0;
  const previewSpacing = previewCharacterSpacing(draft);
  const previewPosition = previewCharacterPosition(draft);

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
          <legend>字符间距与位置</legend>
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
          {(characterSpacingTouched || characterPositionTouched) && error && (
            <p className="work-document-font-dialog-error" role="alert">
              {error}
            </p>
          )}
        </fieldset>
        <section
          className="work-document-font-dialog-preview"
          aria-label="字符间距和位置预览"
        >
          <span>预览</span>
          <output
            style={{
              fontFamily: source.fontFamily ?? undefined,
              fontSize: source.fontSize ?? undefined,
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

function fontDialogDescription(source: DocumentFontDialogSource): string {
  return source.selectedCharacters
    ? `精确设置当前选中内容的原生字符间距和位置（${source.selectedCharacters} 个字符）。`
    : '设置当前位置后续输入文字的原生字符间距和位置。';
}
