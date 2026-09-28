import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
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
import { useOfficeMessages } from './office-messages-context';

type ProofingStateDraft = 'check' | 'inherit' | 'mixed' | 'skip';
type ProofingLanguageDraftKey = 'bidi' | 'eastAsia' | 'latin';

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
  const messages = useOfficeMessages();
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
  const error = proofingDialogError(messages, draft, touched);
  const patch = proofingDialogPatch(source, draft, touched);
  const hasChanges = Object.keys(patch).length > 0;

  const resetDraft = () => {
    setDraft({
      latin: source.latin.value ?? '',
      eastAsia: source.eastAsia.value ?? '',
      bidi: source.bidi.value ?? '',
      noProof: proofingStateDraft(source),
    });
    setTouched({ latin: false, eastAsia: false, bidi: false, noProof: false });
  };

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (!hasChanges || error) return;
    if (onApply(patch)) onClose();
  };

  return (
    <Dialog
      title={officeMessage(messages, 'document.proofing.title')}
      description={
        source.selectedCharacters
          ? officeMessage(
              messages,
              'document.proofing.description.selection',
              { count: String(source.selectedCharacters) },
            )
          : officeMessage(messages, 'document.proofing.description.caret')
      }
      className="work-document-proofing-dialog"
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
            {officeMessage(messages, 'document.proofing.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!hasChanges || Boolean(error)}
          >
            {officeMessage(messages, 'document.proofing.apply')}
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset>
          <legend>
            {officeMessage(messages, 'document.proofing.legend')}
          </legend>
          <ProofingLanguageField
            label={officeMessage(messages, 'document.proofing.latin')}
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
            label={officeMessage(messages, 'document.proofing.eastAsia')}
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
            label={officeMessage(messages, 'document.proofing.complex')}
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
            {officeMessage(messages, 'document.proofing.behavior')}
          </span>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.proofing.behaviorAria',
            )}
            value={draft.noProof}
            options={proofingStateOptions(messages)}
            onValueChange={(noProof) => {
              setDraft((current) => ({ ...current, noProof }));
              setTouched((current) => ({ ...current, noProof: true }));
            }}
          />
          {source.noProof.mixed && !touched.noProof && (
            <p role="status">
              {officeMessage(
                messages,
                'document.proofing.behavior.mixedStatus',
              )}
            </p>
          )}
        </div>

        <p className="work-document-proofing-dialog-help">
          {officeMessage(messages, 'document.proofing.help')}
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
  const messages = useOfficeMessages();
  return (
    <div className="work-document-proofing-dialog-language">
      <label htmlFor={`work-document-proofing-${languageKey}`}>{label}</label>
      <div>
        <OfficeTextField
          id={`work-document-proofing-${languageKey}`}
          aria-label={officeMessage(messages, 'document.proofing.langAria', {
            label,
          })}
          aria-invalid={error || undefined}
          data-autofocus={initialFocus ? 'true' : undefined}
          list={list}
          value={value}
          placeholder={officeMessage(
            messages,
            mixed
              ? 'document.proofing.behavior.mixed'
              : 'document.proofing.behavior.inherit',
          )}
          spellCheck={false}
          onChange={(event) => onChange(event.currentTarget.value)}
        />
        <Button
          tone="quiet"
          type="button"
          aria-label={officeMessage(
            messages,
            'document.proofing.inheritAria',
            { label },
          )}
          onClick={() => onChange('')}
        >
          {officeMessage(messages, 'document.proofing.inherit')}
        </Button>
      </div>
      {mixed && (
        <p role="status">
          {officeMessage(messages, 'document.proofing.mixedHint', { label })}
        </p>
      )}
    </div>
  );
}

function proofingStateOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<ProofingStateDraft>[] {
  return [
    {
      value: 'mixed',
      label: officeMessage(messages, 'document.proofing.behavior.mixed'),
      disabled: true,
    },
    {
      value: 'inherit',
      label: officeMessage(messages, 'document.proofing.behavior.inherit'),
    },
    {
      value: 'check',
      label: officeMessage(messages, 'document.proofing.behavior.check'),
    },
    {
      value: 'skip',
      label: officeMessage(messages, 'document.proofing.behavior.skip'),
    },
  ];
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
  messages: OfficeMessageCatalog,
  draft: Record<ProofingLanguageDraftKey, string>,
  touched: Record<ProofingLanguageDraftKey | 'noProof', boolean>,
): string | null {
  for (const key of ['latin', 'eastAsia', 'bidi'] as const) {
    if (touched[key] && !validLanguageDraft(draft[key])) {
      return officeMessage(messages, 'document.proofing.error.bcp47');
    }
  }
  return null;
}

function validLanguageDraft(value: string): boolean {
  return value === '' || normalizeDocumentLanguageTag(value) !== null;
}
