import type { Editor } from '@tiptap/core';
import { BookMarked, Plus } from 'lucide-react';
import { type FocusEvent, useEffect, useRef, useState } from 'react';
import { Button } from '../../../design-system/primitives';
import {
  createDocumentBibliography,
  documentCitationStyle,
  documentCitationStyleDetails,
  isValidDocumentCitationTag,
} from '../work-document-citations';
import { createWorkId } from '../work-templates';
import type {
  WorkDocumentBibliography,
  WorkDocumentCitationPerson,
  WorkDocumentCitationSource,
  WorkDocumentCitationStyle,
  WorkDocumentContent,
} from '../work-types';
import {
  type CitationSourceDraft,
  type CitationSourceValidationField,
  DocumentCitationSourceForm,
} from './document-citation-source-form';
import { OfficeSelect, useOfficeDialog } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { DocumentTaskPane } from './document-task-pane';
import { officeMessage } from '../../../i18n/office-locale';

export function DocumentCitationsPanel({
  editor,
  content,
  onClose,
  onDirtyChange,
}: {
  editor: Editor;
  content: WorkDocumentContent;
  onClose: () => void;
  onDirtyChange?: (dirty: boolean) => void;
}) {
  const messages = useOfficeMessages();
  const bibliography = content.bibliography ?? createDocumentBibliography();
  const officeDialog = useOfficeDialog();
  const [selectedId, setSelectedId] = useState<string | null>(
    bibliography.sources[0]?.id ?? null,
  );
  const [draft, setDraft] = useState<CitationSourceDraft>(() =>
    sourceDraft(bibliography.sources[0]),
  );
  const [error, setError] = useState<{
    field: CitationSourceValidationField;
    message: string;
  } | null>(null);
  const tagInputRef = useRef<HTMLInputElement>(null);
  const titleInputRef = useRef<HTMLInputElement>(null);
  const draftFocusRef = useRef<HTMLElement | null>(null);
  const selectedSource = bibliography.sources.find(
    (source) => source.id === draft.id,
  );
  const dirty = !sameSourceDraft(draft, sourceDraft(selectedSource));

  useEffect(() => {
    const selected = bibliography.sources.find(
      (source) => source.id === selectedId,
    );
    if (selected) setDraft(sourceDraft(selected));
    else if (selectedId) {
      setSelectedId(bibliography.sources[0]?.id ?? null);
      setDraft(sourceDraft(bibliography.sources[0]));
    }
  }, [bibliography.sources, selectedId]);

  useEffect(() => {
    onDirtyChange?.(dirty);
  }, [dirty, onDirtyChange]);

  useEffect(
    () => () => {
      onDirtyChange?.(false);
    },
    [onDirtyChange],
  );

  const commitBibliography = (
    nextBibliography: WorkDocumentBibliography,
    renamedTag?: { previous: string; next: string },
  ) => {
    editor.commands.setDocumentBibliography(nextBibliography, renamedTag);
  };
  const selectSource = (source: WorkDocumentCitationSource) => {
    setSelectedId(source.id);
    setDraft(sourceDraft(source));
    setError(null);
  };
  const startNewSource = () => {
    setSelectedId(null);
    setDraft(sourceDraft());
    setError(null);
    requestAnimationFrame(() =>
      tagInputRef.current?.focus({ preventScroll: true }),
    );
  };
  const restoreDraftFocusTarget = () => {
    if (draftFocusRef.current?.isConnected) return draftFocusRef.current;
    if (tagInputRef.current?.isConnected) return tagInputRef.current;
    return editor.view.dom;
  };
  const rememberDraftFocus = (event: FocusEvent<HTMLDivElement>) => {
    const target = event.target;
    if (
      target instanceof HTMLElement &&
      target.closest('form') &&
      target.matches(CITATION_DRAFT_CONTROL_SELECTOR)
    ) {
      draftFocusRef.current = target;
    }
  };
  const continueAfterDiscard = async (action: () => void) => {
    if (
      dirty &&
      !(await officeDialog.confirm({
        title: officeMessage(messages, 'document.citation.discard.title'),
        description: officeMessage(
          messages,
          'document.citation.discard.description',
        ),
        confirmLabel: officeMessage(
          messages,
          'document.citation.discard.confirm',
        ),
        confirmTone: 'danger',
        restoreFocusTarget: restoreDraftFocusTarget,
      }))
    ) {
      return;
    }
    action();
  };
  const reportValidationError = (
    field: CitationSourceValidationField,
    message: string,
  ) => {
    setError({ field, message });
    requestAnimationFrame(() => {
      const target =
        field === 'tag' ? tagInputRef.current : titleInputRef.current;
      target?.focus({ preventScroll: true });
      target?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    });
  };
  const saveSource = () => {
    const tag = draft.tag.trim();
    if (!isValidDocumentCitationTag(tag)) {
      reportValidationError(
        'tag',
        officeMessage(messages, 'document.citation.error.tagFormat'),
      );
      return;
    }
    if (!draft.title.trim()) {
      reportValidationError(
        'title',
        officeMessage(messages, 'document.citation.error.titleRequired'),
      );
      return;
    }
    if (
      bibliography.sources.some(
        (source) =>
          source.id !== draft.id &&
          source.tag.toLowerCase() === tag.toLowerCase(),
      )
    ) {
      reportValidationError(
        'tag',
        officeMessage(messages, 'document.citation.error.tagDuplicate'),
      );
      return;
    }
    const existing = bibliography.sources.find(
      (source) => source.id === draft.id,
    );
    const authorPeople = parseAuthors(draft.authors);
    const author =
      draft.corporateAuthor.trim() || authorPeople.length
        ? {
            corporate: draft.corporateAuthor.trim() || undefined,
            people: draft.corporateAuthor.trim() ? undefined : authorPeople,
          }
        : undefined;
    const contributors = { ...(existing?.contributors ?? {}) };
    if (author) contributors.Author = author;
    else delete contributors.Author;
    const saved: WorkDocumentCitationSource = {
      id: draft.id ?? createWorkId('source'),
      tag,
      sourceType: draft.sourceType || 'Misc',
      guid: existing?.guid,
      title: draft.title.trim(),
      year: optionalValue(draft.year),
      contributors: Object.keys(contributors).length ? contributors : undefined,
      publisher: optionalValue(draft.publisher),
      city: optionalValue(draft.city),
      journalName: optionalValue(draft.journalName),
      volume: optionalValue(draft.volume),
      issue: optionalValue(draft.issue),
      pages: optionalValue(draft.pages),
      url: optionalValue(draft.url),
      standardNumber: optionalValue(draft.standardNumber),
      conferenceName: optionalValue(draft.conferenceName),
      institution: optionalValue(draft.institution),
      additionalFields: existing?.additionalFields,
    };
    const sources = draft.id
      ? bibliography.sources.map((source) =>
          source.id === draft.id ? saved : source,
        )
      : [...bibliography.sources, saved];
    commitBibliography(
      { ...bibliography, sources },
      existing && existing.tag !== saved.tag
        ? { previous: existing.tag, next: saved.tag }
        : undefined,
    );
    setSelectedId(saved.id);
    setDraft(sourceDraft(saved));
    setError(null);
  };
  const deleteSource = async () => {
    if (!draft.id) {
      startNewSource();
      return;
    }
    const sourceTitle =
      selectedSource?.title ||
      officeMessage(messages, 'document.citation.delete.fallbackTitle');
    const confirmed = await officeDialog.confirm({
      title: officeMessage(messages, 'document.citation.delete.title'),
      description: officeMessage(
        messages,
        'document.citation.delete.description',
        { title: sourceTitle },
      ),
      confirmLabel: officeMessage(
        messages,
        'document.citation.delete.confirm',
      ),
      confirmTone: 'danger',
      restoreFocusTarget: dirty ? restoreDraftFocusTarget : undefined,
    });
    if (!confirmed) return;
    const sources = bibliography.sources.filter(
      (source) => source.id !== draft.id,
    );
    commitBibliography({ ...bibliography, sources });
    const next = sources[0];
    setSelectedId(next?.id ?? null);
    setDraft(sourceDraft(next));
    setError(null);
  };
  const changeStyle = (style: WorkDocumentCitationStyle) => {
    const details = documentCitationStyleDetails(style);
    commitBibliography({
      ...bibliography,
      style,
      styleName: details.name,
      selectedStyle: details.selectedStyle,
    });
  };
  return (
    <>
      <DocumentTaskPane
        className="work-document-citations-panel"
        title={officeMessage(messages, 'document.citation.panel.title')}
        description={officeMessage(
          messages,
          dirty
            ? 'document.citation.panel.descriptionDirty'
            : 'document.citation.panel.description',
          { count: String(bibliography.sources.length) },
        )}
        closeLabel={officeMessage(messages, 'document.citation.panel.close')}
        onClose={onClose}
      >
        <div className="work-document-citation-actions">
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'document.citation.panel.style')}
            </span>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'document.citation.panel.styleAria',
              )}
              value={documentCitationStyle(bibliography.style)}
              options={[
                { value: 'apa', label: 'APA' },
                { value: 'mla', label: 'MLA' },
                { value: 'chicago', label: 'Chicago' },
                { value: 'ieee', label: 'IEEE' },
              ]}
              onValueChange={changeStyle}
            />
          </div>
          <Button
            tone="secondary"
            aria-label={officeMessage(
              messages,
              'document.citation.panel.insertBibliography',
            )}
            disabled={!bibliography.sources.length}
            onClick={() =>
              editor
                .chain()
                .focus()
                .insertDocumentBibliography(bibliography)
                .run()
            }
          >
            <BookMarked size={13} />
            {officeMessage(
              messages,
              'document.citation.panel.insertBibliography',
            )}
          </Button>
        </div>
        <div
          className="work-document-citation-manager"
          onFocusCapture={rememberDraftFocus}
        >
          {bibliography.sources.length > 0 && (
            <aside
              aria-label={officeMessage(
                messages,
                'document.citation.panel.listAria',
              )}
            >
              <div className="work-document-citation-list-heading">
                <strong>
                  {officeMessage(
                    messages,
                    'document.citation.panel.listHeading',
                  )}
                </strong>
                <Button
                  className="create"
                  size="compact"
                  tone="quiet"
                  onClick={() => void continueAfterDiscard(startNewSource)}
                >
                  <Plus size={13} />
                  {officeMessage(messages, 'document.citation.panel.create')}
                </Button>
              </div>
              <div className="work-document-citation-source-list">
                {bibliography.sources.map((source) => (
                  <button
                    type="button"
                    className={source.id === selectedId ? 'active' : ''}
                    aria-current={source.id === selectedId}
                    key={source.id}
                    onClick={() =>
                      void continueAfterDiscard(() => selectSource(source))
                    }
                  >
                    <strong>
                      {source.title ||
                        officeMessage(
                          messages,
                          'document.citation.panel.untitled',
                        )}
                    </strong>
                    <span>
                      {officeMessage(
                        messages,
                        'document.citation.panel.sourceMeta',
                        {
                          tag: source.tag,
                          year:
                            source.year ||
                            officeMessage(
                              messages,
                              'document.citation.panel.noYear',
                            ),
                        },
                      )}
                    </span>
                  </button>
                ))}
              </div>
            </aside>
          )}
          <DocumentCitationSourceForm
            draft={draft}
            dirty={dirty}
            error={error?.message ?? ''}
            errorField={error?.field}
            tagInputRef={tagInputRef}
            titleInputRef={titleInputRef}
            onDraftChange={(nextDraft) => {
              setDraft(nextDraft);
              setError((current) =>
                current && nextDraft[current.field] !== draft[current.field]
                  ? null
                  : current,
              );
            }}
            onSave={saveSource}
            onInsert={() => {
              if (!selectedSource || dirty) return;
              editor
                .chain()
                .focus()
                .insertDocumentCitation(selectedSource, bibliography)
                .run();
            }}
            onDelete={() => void deleteSource()}
          />
        </div>
      </DocumentTaskPane>
      {officeDialog.dialog}
    </>
  );
}

const CITATION_DRAFT_CONTROL_SELECTOR = [
  'input:not([type="hidden"]):not(:disabled)',
  'textarea:not(:disabled)',
  '[role="combobox"]:not([aria-disabled="true"])',
].join(', ');

function sourceDraft(source?: WorkDocumentCitationSource): CitationSourceDraft {
  const author = source?.contributors?.Author;
  return {
    id: source?.id,
    tag: source?.tag ?? '',
    sourceType: source?.sourceType ?? 'Book',
    title: source?.title ?? '',
    year: source?.year ?? '',
    authors: formatAuthors(author?.people ?? []),
    corporateAuthor: author?.corporate ?? '',
    publisher: source?.publisher ?? '',
    city: source?.city ?? '',
    journalName: source?.journalName ?? '',
    volume: source?.volume ?? '',
    issue: source?.issue ?? '',
    pages: source?.pages ?? '',
    url: source?.url ?? '',
    standardNumber: source?.standardNumber ?? '',
    conferenceName: source?.conferenceName ?? '',
    institution: source?.institution ?? '',
  };
}

function formatAuthors(people: WorkDocumentCitationPerson[]): string {
  return people
    .map((person) => {
      const given = [person.first, person.middle].filter(Boolean).join(' ');
      return [person.last, given].filter(Boolean).join(', ');
    })
    .join('\n');
}

function parseAuthors(value: string): WorkDocumentCitationPerson[] {
  return value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [lastPart, givenPart] = line.split(',', 2);
      if (givenPart !== undefined) {
        const given = givenPart.trim().split(/\s+/).filter(Boolean);
        return {
          first: given.shift() ?? '',
          middle: given.join(' ') || undefined,
          last: lastPart.trim(),
        };
      }
      const parts = line.split(/\s+/);
      const last = parts.pop() ?? '';
      return {
        first: parts.shift() ?? '',
        middle: parts.join(' ') || undefined,
        last,
      };
    });
}

function optionalValue(value: string): string | undefined {
  return value.trim() || undefined;
}

function sameSourceDraft(
  left: CitationSourceDraft,
  right: CitationSourceDraft,
): boolean {
  return (Object.keys(left) as Array<keyof CitationSourceDraft>).every(
    (key) => left[key] === right[key],
  );
}
