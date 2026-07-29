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
  DocumentCitationSourceForm,
} from './document-citation-source-form';
import { OfficeSelect, useOfficeDialog } from './office-controls';
import { DocumentTaskPane } from './document-task-pane';

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
  const bibliography = content.bibliography ?? createDocumentBibliography();
  const officeDialog = useOfficeDialog();
  const [selectedId, setSelectedId] = useState<string | null>(
    bibliography.sources[0]?.id ?? null,
  );
  const [draft, setDraft] = useState<CitationSourceDraft>(() =>
    sourceDraft(bibliography.sources[0]),
  );
  const [error, setError] = useState('');
  const tagInputRef = useRef<HTMLInputElement>(null);
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
    setError('');
  };
  const startNewSource = () => {
    setSelectedId(null);
    setDraft(sourceDraft());
    setError('');
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
        title: '放弃未保存的更改？',
        description: '当前文献尚未保存。',
        confirmLabel: '放弃更改',
        confirmTone: 'danger',
        restoreFocusTarget: restoreDraftFocusTarget,
      }))
    ) {
      return;
    }
    action();
  };
  const saveSource = () => {
    const tag = draft.tag.trim();
    if (!isValidDocumentCitationTag(tag)) {
      setError('简称只能使用字母、数字、下划线及 . : + -，长度不超过 80。');
      return;
    }
    if (!draft.title.trim()) {
      setError('请输入文献标题。');
      return;
    }
    if (
      bibliography.sources.some(
        (source) =>
          source.id !== draft.id &&
          source.tag.toLowerCase() === tag.toLowerCase(),
      )
    ) {
      setError('已经存在相同的简称。');
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
    setError('');
  };
  const deleteSource = async () => {
    if (!draft.id) {
      startNewSource();
      return;
    }
    const sourceTitle = selectedSource?.title || '当前文献';
    const confirmed = await officeDialog.confirm({
      title: '删除文献？',
      description: `“${sourceTitle}”将从文献库中删除，文档中的引用可能无法识别。`,
      confirmLabel: '删除',
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
    setError('');
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
        title="文献库"
        description={`${bibliography.sources.length} 条文献${dirty ? ' · 有未保存更改' : ''}`}
        closeLabel="关闭文献库"
        onClose={onClose}
      >
        <div className="work-document-citation-actions">
          <div className="work-office-field">
            <span>样式</span>
            <OfficeSelect
              ariaLabel="引文样式"
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
            aria-label="插入参考文献"
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
            插入参考文献
          </Button>
        </div>
        <div
          className="work-document-citation-manager"
          onFocusCapture={rememberDraftFocus}
        >
          {bibliography.sources.length > 0 && (
            <aside aria-label="文献列表">
              <div className="work-document-citation-list-heading">
                <strong>文献</strong>
                <Button
                  className="create"
                  size="compact"
                  tone="quiet"
                  onClick={() => void continueAfterDiscard(startNewSource)}
                >
                  <Plus size={13} />
                  新建
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
                    <strong>{source.title || '未命名文献'}</strong>
                    <span>
                      {source.tag} · {source.year || '无年份'}
                    </span>
                  </button>
                ))}
              </div>
            </aside>
          )}
          <DocumentCitationSourceForm
            draft={draft}
            dirty={dirty}
            error={error}
            tagInputRef={tagInputRef}
            onDraftChange={(nextDraft) => {
              setDraft(nextDraft);
              setError('');
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
