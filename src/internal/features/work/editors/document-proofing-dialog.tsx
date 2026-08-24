import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { normalizeDocumentLanguageTag } from '../work-document-proofing';
import type {
  DocumentProofingDialogPatch,
  DocumentProofingDialogSource,
} from './document-proofing-dialog-model';
import {
  OfficeSelect,
  OfficeTextField,
  type OfficeSelectOption,
} from './office-controls';

type ProofingStateDraft = 'check' | 'inherit' | 'mixed' | 'skip';
type ProofingLanguageDraftKey = 'bidi' | 'eastAsia' | 'latin';

const proofingStateOptions: readonly OfficeSelectOption<ProofingStateDraft>[] =
  [
    { value: 'mixed', label: '混合（保持不变）', disabled: true },
    { value: 'inherit', label: '跟随样式' },
    { value: 'check', label: '检查拼写和语法' },
    { value: 'skip', label: '不检查拼写或语法' },
  ];

const proofingLanguageSuggestions = [
  'en-US',
  'en-GB',
  'zh-CN',
  'zh-TW',
  'ja-JP',
  'ko-KR',
  'ar-SA',
  'he-IL',
  'fr-FR',
  'de-DE',
  'es-ES',
  'x-none',
] as const;

export function DocumentProofingDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  source: DocumentProofingDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (patch: DocumentProofingDialogPatch) => boolean;
  onClose: () => void;
}) {
  const [draft, setDraft] = useState(() => ({
    latin: source.latin.value ?? '',
    eastAsia: source.eastAsia.value ?? '',
    bidi: source.bidi.value ?? '',
    noProof: proofingStateDraft(source),
  }));
  const [touched, setTouched] = useState<
    Record<ProofingLanguageDraftKey | 'noProof', boolean>
  >({ latin: false, eastAsia: false, bidi: false, noProof: false });
  const formId = useId();
  const datalistId = useId();
  const error = proofingDialogError(draft, touched);
  const patch = proofingDialogPatch(source, draft, touched);
  const hasChanges = Object.keys(patch).length > 0;

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (!hasChanges || error) return;
    if (onApply(patch)) onClose();
  };

  return (
    <Dialog
      title="设置校对语言"
      description={
        source.selectedCharacters
          ? `为选中的 ${source.selectedCharacters} 个字符分别设置拉丁、东亚和双向文字校对语言。`
          : '设置当前位置后续输入文字的校对语言。'
      }
      className="work-document-proofing-dialog"
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
        <fieldset>
          <legend>按文字系统设置语言</legend>
          <ProofingLanguageField
            label="拉丁文字"
            languageKey="latin"
            value={draft.latin}
            mixed={source.latin.mixed && !touched.latin}
            list={datalistId}
            initialFocus
            error={touched.latin && !validLanguageDraft(draft.latin)}
            onChange={(value) => {
              setDraft((current) => ({ ...current, latin: value }));
              setTouched((current) => ({ ...current, latin: true }));
            }}
          />
          <ProofingLanguageField
            label="东亚文字"
            languageKey="eastAsia"
            value={draft.eastAsia}
            mixed={source.eastAsia.mixed && !touched.eastAsia}
            list={datalistId}
            error={touched.eastAsia && !validLanguageDraft(draft.eastAsia)}
            onChange={(value) => {
              setDraft((current) => ({ ...current, eastAsia: value }));
              setTouched((current) => ({ ...current, eastAsia: true }));
            }}
          />
          <ProofingLanguageField
            label="双向文字"
            languageKey="bidi"
            value={draft.bidi}
            mixed={source.bidi.mixed && !touched.bidi}
            list={datalistId}
            error={touched.bidi && !validLanguageDraft(draft.bidi)}
            onChange={(value) => {
              setDraft((current) => ({ ...current, bidi: value }));
              setTouched((current) => ({ ...current, bidi: true }));
            }}
          />
          <datalist id={datalistId}>
            {proofingLanguageSuggestions.map((language) => (
              <option key={language} value={language} />
            ))}
          </datalist>
        </fieldset>

        <div className="work-document-proofing-dialog-state">
          <span
            className="work-document-proofing-dialog-state-label"
            aria-hidden="true"
          >
            校对行为
          </span>
          <OfficeSelect
            ariaLabel="校对行为"
            value={draft.noProof}
            options={proofingStateOptions}
            onValueChange={(noProof) => {
              setDraft((current) => ({ ...current, noProof }));
              setTouched((current) => ({ ...current, noProof: true }));
            }}
          />
          {source.noProof.mixed && !touched.noProof && (
            <p role="status">当前选区包含不同的校对行为，保持不变。</p>
          )}
        </div>

        <p className="work-document-proofing-dialog-help">
          使用 BCP 47 语言标记，例如 <code>en-US</code>、<code>zh-CN</code> 或{' '}
          <code>ar-SA</code>。留空或选择“跟随样式”会移除直接格式。
        </p>
        {error && (
          <p className="work-document-proofing-dialog-error" role="alert">
            {error}
          </p>
        )}
      </form>
    </Dialog>
  );
}

function ProofingLanguageField({
  label,
  languageKey,
  value,
  mixed,
  list,
  initialFocus = false,
  error,
  onChange,
}: {
  label: string;
  languageKey: ProofingLanguageDraftKey;
  value: string;
  mixed: boolean;
  list: string;
  initialFocus?: boolean;
  error: boolean;
  onChange: (value: string) => void;
}) {
  return (
    <div className="work-document-proofing-dialog-language">
      <label htmlFor={`work-document-proofing-${languageKey}`}>{label}</label>
      <div>
        <OfficeTextField
          id={`work-document-proofing-${languageKey}`}
          aria-label={`${label}校对语言`}
          aria-invalid={error || undefined}
          data-autofocus={initialFocus ? 'true' : undefined}
          list={list}
          value={value}
          placeholder={mixed ? '混合（保持不变）' : '跟随样式'}
          spellCheck={false}
          onChange={(event) => onChange(event.currentTarget.value)}
        />
        <Button
          tone="quiet"
          type="button"
          aria-label={`${label}跟随样式`}
          onClick={() => onChange('')}
        >
          跟随样式
        </Button>
      </div>
      {mixed && (
        <p role="status">
          当前选区包含不同的{label}校对语言，输入后才会统一修改。
        </p>
      )}
    </div>
  );
}

function proofingStateDraft(
  source: DocumentProofingDialogSource,
): ProofingStateDraft {
  if (source.noProof.mixed) return 'mixed';
  if (source.noProof.value === null) return 'inherit';
  return source.noProof.value ? 'skip' : 'check';
}

function proofingDialogPatch(
  source: DocumentProofingDialogSource,
  draft: Record<ProofingLanguageDraftKey, string> & {
    noProof: ProofingStateDraft;
  },
  touched: Record<ProofingLanguageDraftKey | 'noProof', boolean>,
): DocumentProofingDialogPatch {
  const languages: NonNullable<DocumentProofingDialogPatch['languages']> = {};
  for (const key of ['latin', 'eastAsia', 'bidi'] as const) {
    if (!touched[key] || !validLanguageDraft(draft[key])) continue;
    const value = draft[key] || null;
    if (source[key].mixed || source[key].value !== value) {
      languages[key] = value;
    }
  }
  const patch: DocumentProofingDialogPatch = {};
  if (Object.keys(languages).length) patch.languages = languages;
  if (touched.noProof && draft.noProof !== 'mixed') {
    const noProof =
      draft.noProof === 'inherit' ? null : draft.noProof === 'skip';
    if (source.noProof.mixed || source.noProof.value !== noProof) {
      patch.noProof = noProof;
    }
  }
  return patch;
}

function proofingDialogError(
  draft: Record<ProofingLanguageDraftKey, string>,
  touched: Record<ProofingLanguageDraftKey | 'noProof', boolean>,
): string | null {
  for (const key of ['latin', 'eastAsia', 'bidi'] as const) {
    if (touched[key] && !validLanguageDraft(draft[key])) {
      return '请输入有效的 BCP 47 语言标记；不能包含空格、下划线或控制字符。';
    }
  }
  return null;
}

function validLanguageDraft(value: string): boolean {
  return value === '' || normalizeDocumentLanguageTag(value) !== null;
}
