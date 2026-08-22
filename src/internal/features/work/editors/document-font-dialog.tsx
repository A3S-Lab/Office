import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeNumberField, OfficeSelect } from './office-controls';
import {
  createDocumentFontDialogDraft,
  documentFontDialogDraftError,
  documentFontDialogPatch,
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
  const formId = useId();
  const error = documentFontDialogDraftError(draft);
  const patch = documentFontDialogPatch(source, draft, characterSpacingTouched);
  const hasChanges = patch.characterSpacingTwips !== undefined;
  const previewSpacing = previewCharacterSpacing(draft);

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
          <legend>字符间距</legend>
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
          {source.characterSpacing.mixed && !characterSpacingTouched && (
            <p className="work-document-font-dialog-mixed" role="status">
              当前选区包含多种字符间距。选择一种间距后才会统一修改。
            </p>
          )}
          {characterSpacingTouched && error && (
            <p className="work-document-font-dialog-error" role="alert">
              {error}
            </p>
          )}
        </fieldset>
        <section
          className="work-document-font-dialog-preview"
          aria-label="字符间距预览"
        >
          <span>预览</span>
          <output
            style={{
              fontFamily: source.fontFamily ?? undefined,
              fontSize: source.fontSize ?? undefined,
              letterSpacing: `${previewSpacing}pt`,
            }}
          >
            {source.previewText}
          </output>
        </section>
      </form>
    </Dialog>
  );
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
    ? `精确设置当前选中内容的原生字符间距（${source.selectedCharacters} 个字符）。`
    : '设置当前位置后续输入文字的原生字符间距。';
}
