import { Quote, Trash2 } from 'lucide-react';
import { type Ref, useId } from 'react';
import { Button } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type {
  OfficeMessageCatalog,
  OfficeMessageKey,
} from '../../../i18n/office-messages';
import {
  OfficeSelect,
  OfficeTextArea,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface CitationSourceDraft {
  id?: string;
  tag: string;
  sourceType: string;
  title: string;
  year: string;
  authors: string;
  corporateAuthor: string;
  publisher: string;
  city: string;
  journalName: string;
  volume: string;
  issue: string;
  pages: string;
  url: string;
  standardNumber: string;
  conferenceName: string;
  institution: string;
}

export type CitationSourceValidationField = 'tag' | 'title';

const SOURCE_TYPE_KEYS = [
  ['Book', 'document.citation.source.Book'],
  ['BookSection', 'document.citation.source.BookSection'],
  ['JournalArticle', 'document.citation.source.JournalArticle'],
  ['ArticleInAPeriodical', 'document.citation.source.ArticleInAPeriodical'],
  ['ConferenceProceedings', 'document.citation.source.ConferenceProceedings'],
  ['Report', 'document.citation.source.Report'],
  ['InternetSite', 'document.citation.source.InternetSite'],
  [
    'DocumentFromInternetSite',
    'document.citation.source.DocumentFromInternetSite',
  ],
  ['ElectronicSource', 'document.citation.source.ElectronicSource'],
  ['Misc', 'document.citation.source.Misc'],
] as const satisfies ReadonlyArray<readonly [string, OfficeMessageKey]>;

export function DocumentCitationSourceForm({
  draft,
  dirty,
  error,
  errorField,
  tagInputRef,
  titleInputRef,
  onDraftChange,
  onSave,
  onInsert,
  onDelete,
}: {
  draft: CitationSourceDraft;
  dirty: boolean;
  error: string;
  errorField?: CitationSourceValidationField;
  tagInputRef?: Ref<HTMLInputElement>;
  titleInputRef?: Ref<HTMLInputElement>;
  onDraftChange: (draft: CitationSourceDraft) => void;
  onSave: () => void;
  onInsert: () => void;
  onDelete: () => void;
}) {
  const messages = useOfficeMessages();
  const validationErrorId = useId();
  const saved = Boolean(draft.id);
  const knownSourceType = SOURCE_TYPE_KEYS.some(
    ([value]) => value === draft.sourceType,
  );
  const update = <Key extends keyof CitationSourceDraft>(
    key: Key,
    value: CitationSourceDraft[Key],
  ) => onDraftChange({ ...draft, [key]: value });
  const tagError = errorField === 'tag' ? error : '';
  const titleError = errorField === 'title' ? error : '';
  const formTitle = officeMessage(
    messages,
    saved ? 'document.citation.form.edit' : 'document.citation.form.create',
  );

  return (
    <form
      aria-label={formTitle}
      onSubmit={(event) => {
        event.preventDefault();
        onSave();
      }}
    >
      <div className="work-document-citation-form-heading wide">
        <strong>{formTitle}</strong>
        <span>
          {officeMessage(messages, 'document.citation.form.requiredHint')}
        </span>
      </div>
      <div className="work-office-field">
        <span>{officeMessage(messages, 'document.citation.form.tag')}</span>
        <OfficeTextField
          ref={tagInputRef}
          aria-label={officeMessage(messages, 'document.citation.form.tagAria')}
          aria-describedby={tagError ? validationErrorId : undefined}
          aria-invalid={tagError ? true : undefined}
          value={draft.tag}
          maxLength={80}
          placeholder={officeMessage(
            messages,
            'document.citation.form.tagPlaceholder',
          )}
          onChange={(event) => update('tag', event.target.value)}
        />
        {tagError && (
          <span
            id={validationErrorId}
            className="work-office-dialog-field-error"
            role="alert"
          >
            {tagError}
          </span>
        )}
      </div>
      <div className="work-office-field">
        <span>{officeMessage(messages, 'document.citation.form.type')}</span>
        <OfficeSelect
          ariaLabel={officeMessage(
            messages,
            'document.citation.form.typeAria',
          )}
          value={draft.sourceType}
          options={[
            ...(!knownSourceType && draft.sourceType
              ? [
                  {
                    value: draft.sourceType,
                    label: officeMessage(
                      messages,
                      'document.citation.form.rawType',
                      { type: draft.sourceType },
                    ),
                  },
                ]
              : []),
            ...citationSourceTypeOptions(messages),
          ]}
          onValueChange={(sourceType) => update('sourceType', sourceType)}
        />
      </div>
      <div className="work-office-field wide">
        <span>{officeMessage(messages, 'document.citation.form.title')}</span>
        <OfficeTextField
          ref={titleInputRef}
          aria-label={officeMessage(
            messages,
            'document.citation.form.titleAria',
          )}
          aria-describedby={titleError ? validationErrorId : undefined}
          aria-invalid={titleError ? true : undefined}
          value={draft.title}
          onChange={(event) => update('title', event.target.value)}
        />
        {titleError && (
          <span
            id={validationErrorId}
            className="work-office-dialog-field-error"
            role="alert"
          >
            {titleError}
          </span>
        )}
      </div>
      <div className="work-office-field">
        <span>{officeMessage(messages, 'document.citation.form.year')}</span>
        <OfficeTextField
          aria-label={officeMessage(
            messages,
            'document.citation.form.yearAria',
          )}
          value={draft.year}
          inputMode="numeric"
          placeholder="2026"
          onChange={(event) => update('year', event.target.value)}
        />
      </div>
      <div className="work-office-field">
        <span>
          {officeMessage(messages, 'document.citation.form.corporateAuthor')}
        </span>
        <OfficeTextField
          aria-label={officeMessage(
            messages,
            'document.citation.form.corporateAuthorAria',
          )}
          value={draft.corporateAuthor}
          placeholder={officeMessage(
            messages,
            'document.citation.form.corporateAuthorPlaceholder',
          )}
          onChange={(event) => update('corporateAuthor', event.target.value)}
        />
      </div>
      <div className="work-office-field wide">
        <span>{officeMessage(messages, 'document.citation.form.authors')}</span>
        <OfficeTextArea
          aria-label={officeMessage(
            messages,
            'document.citation.form.authorsAria',
          )}
          value={draft.authors}
          placeholder={officeMessage(
            messages,
            'document.citation.form.authorsPlaceholder',
          )}
          onChange={(event) => update('authors', event.target.value)}
        />
      </div>
      <details className="work-document-citation-more-fields wide">
        <summary>
          {officeMessage(messages, 'document.citation.form.more')}
        </summary>
        <div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.form.publisher')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.publisherAria',
              )}
              value={draft.publisher}
              onChange={(event) => update('publisher', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.form.city')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.cityAria',
              )}
              value={draft.city}
              onChange={(event) => update('city', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.form.journal')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.journalAria',
              )}
              value={draft.journalName}
              onChange={(event) => update('journalName', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.form.volumeIssue')}
            </span>
            <span className="paired">
              <OfficeTextField
                aria-label={officeMessage(
                  messages,
                  'document.citation.form.volumeAria',
                )}
                value={draft.volume}
                placeholder={officeMessage(
                  messages,
                  'document.citation.form.volumePlaceholder',
                )}
                onChange={(event) => update('volume', event.target.value)}
              />
              <OfficeTextField
                aria-label={officeMessage(
                  messages,
                  'document.citation.form.issueAria',
                )}
                value={draft.issue}
                placeholder={officeMessage(
                  messages,
                  'document.citation.form.issuePlaceholder',
                )}
                onChange={(event) => update('issue', event.target.value)}
              />
            </span>
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.form.pages')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.pagesAria',
              )}
              value={draft.pages}
              placeholder="12–28"
              onChange={(event) => update('pages', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(
                messages,
                'document.citation.form.standardNumber',
              )}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.standardNumberAria',
              )}
              value={draft.standardNumber}
              onChange={(event) => update('standardNumber', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.form.conference')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.conferenceAria',
              )}
              value={draft.conferenceName}
              onChange={(event) => update('conferenceName', event.target.value)}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.form.institution')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.institutionAria',
              )}
              value={draft.institution}
              onChange={(event) => update('institution', event.target.value)}
            />
          </div>
          <div className="work-office-field wide">
            <span>
              {officeMessage(messages, 'document.citation.form.url')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.citation.form.urlAria',
              )}
              value={draft.url}
              inputMode="url"
              placeholder="https://"
              onChange={(event) => update('url', event.target.value)}
            />
          </div>
        </div>
      </details>
      <div className="actions wide">
        <div className="work-document-citation-form-buttons">
          {saved && (
            <Button
              tone="danger"
              aria-label={officeMessage(
                messages,
                'document.citation.form.deleteAria',
              )}
              onClick={onDelete}
            >
              <Trash2 size={13} />
              {officeMessage(messages, 'document.citation.form.delete')}
            </Button>
          )}
          <span aria-hidden="true" />
          {saved && (
            <Button tone="secondary" disabled={dirty} onClick={onInsert}>
              <Quote size={13} />
              {officeMessage(messages, 'document.citation.form.insert')}
            </Button>
          )}
          <Button
            type="submit"
            tone="primary"
            aria-label={officeMessage(
              messages,
              'document.citation.form.saveAria',
            )}
            disabled={!dirty}
          >
            {officeMessage(messages, 'document.citation.form.save')}
          </Button>
        </div>
      </div>
    </form>
  );
}

function citationSourceTypeOptions(messages: OfficeMessageCatalog) {
  return SOURCE_TYPE_KEYS.map(([value, key]) => ({
    value,
    label: officeMessage(messages, key),
  }));
}
