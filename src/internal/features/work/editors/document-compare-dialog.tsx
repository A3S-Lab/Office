import { FileCheck2, FileUp, GitCompareArrows } from 'lucide-react';
import { type FormEvent, useId, useRef, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type {
  OfficeMessageCatalog,
  OfficeMessageKey,
} from '../../../i18n/office-messages';
import type {
  DocumentComparisonApplyResult,
  DocumentComparisonDiagnostic,
  DocumentComparisonMode,
} from '../work-document-compare';
import { OfficeFileInput, OfficeTextField } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentCompareDialogRequest {
  author: string;
  file: File;
  mode: DocumentComparisonMode;
}

export interface DocumentCompareDialogProps {
  initialMode: DocumentComparisonMode;
  restoreFocusTarget: () => HTMLElement | null;
  onApplied: (result: DocumentComparisonApplyResult) => void;
  onClose: () => void;
  onSubmit: (
    request: DocumentCompareDialogRequest,
  ) => Promise<DocumentComparisonApplyResult>;
}

const DIAGNOSTIC_MESSAGE_KEYS = {
  'changed-complex-structure':
    'document.compare.diagnostic.changed-complex-structure',
  'combine-baseline-mismatch':
    'document.compare.diagnostic.combine-baseline-mismatch',
  'combine-resolution-invalid':
    'document.compare.diagnostic.combine-resolution-invalid',
  'combine-structural-revisions':
    'document.compare.diagnostic.combine-structural-revisions',
  'combine-without-revisions':
    'document.compare.diagnostic.combine-without-revisions',
  'comparison-limit-exceeded':
    'document.compare.diagnostic.comparison-limit-exceeded',
  'current-revisions-present':
    'document.compare.diagnostic.current-revisions-present',
  'empty-structural-change':
    'document.compare.diagnostic.empty-structural-change',
  'invalid-revised-content':
    'document.compare.diagnostic.invalid-revised-content',
  'revised-revisions-present':
    'document.compare.diagnostic.revised-revisions-present',
  'section-layout-mismatch':
    'document.compare.diagnostic.section-layout-mismatch',
  'unsupported-inline-review-state':
    'document.compare.diagnostic.unsupported-inline-review-state',
} as const satisfies Record<
  DocumentComparisonDiagnostic['code'],
  OfficeMessageKey
>;

export function DocumentCompareDialog({
  initialMode,
  restoreFocusTarget,
  onApplied,
  onClose,
  onSubmit,
}: DocumentCompareDialogProps) {
  const messages = useOfficeMessages();
  const defaultAuthor = officeMessage(
    messages,
    'document.compare.author.default',
  );
  const [mode, setMode] = useState(initialMode);
  const [file, setFile] = useState<File | null>(null);
  const [author, setAuthor] = useState(defaultAuthor);
  const [authorTouched, setAuthorTouched] = useState(false);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<DocumentComparisonApplyResult | null>(
    null,
  );
  const fileInputRef = useRef<HTMLInputElement>(null);
  const formId = useId();
  const authorId = useId();
  const submit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!file || busy || (mode === 'compare' && !author.trim())) return;
    setBusy(true);
    setResult(null);
    try {
      const next = await onSubmit({ author: author.trim(), file, mode });
      setResult(next);
      if (next.status === 'applied') onApplied(next);
    } catch {
      setResult({
        status: 'unsupported',
        summary: {
          deletions: 0,
          formatting: 0,
          insertions: 0,
          paragraphFormatting: 0,
        },
        diagnostics: [
          {
            code: 'invalid-revised-content',
            message: 'The selected file could not be imported.',
          },
        ],
      });
    } finally {
      setBusy(false);
    }
  };
  const selectFile = (next: File) => {
    setFile(next);
    setResult(null);
    if (!authorTouched) {
      setAuthor(fileStem(next.name) || defaultAuthor);
    }
  };

  return (
    <Dialog
      title={officeMessage(messages, 'document.compare.title')}
      description={officeMessage(messages, 'document.compare.description')}
      className="work-document-compare-dialog"
      closeDisabled={busy}
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" disabled={busy} onClick={onClose}>
            {officeMessage(messages, 'document.compare.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!file || busy || (mode === 'compare' && !author.trim())}
          >
            {busy
              ? officeMessage(messages, 'document.compare.busy')
              : officeMessage(
                  messages,
                  mode === 'compare'
                    ? 'document.compare.submit.compare'
                    : 'document.compare.submit.combine',
                )}
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-document-compare-mode">
          <legend>
            {officeMessage(messages, 'document.compare.mode.legend')}
          </legend>
          <label>
            <input
              type="radio"
              name="document-comparison-mode"
              value="compare"
              checked={mode === 'compare'}
              onChange={() => {
                setMode('compare');
                setResult(null);
              }}
            />
            <span>
              <GitCompareArrows size={18} aria-hidden="true" />
              <strong>
                {officeMessage(messages, 'document.compare.mode.compare')}
              </strong>
              <small>
                {officeMessage(messages, 'document.compare.mode.compareHint')}
              </small>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="document-comparison-mode"
              value="combine"
              checked={mode === 'combine'}
              onChange={() => {
                setMode('combine');
                setResult(null);
              }}
            />
            <span>
              <FileCheck2 size={18} aria-hidden="true" />
              <strong>
                {officeMessage(messages, 'document.compare.mode.combine')}
              </strong>
              <small>
                {officeMessage(messages, 'document.compare.mode.combineHint')}
              </small>
            </span>
          </label>
        </fieldset>

        <OfficeFileInput
          ref={fileInputRef}
          accept=".docx,.html,.htm,.txt,application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/html,text/plain"
          aria-label={officeMessage(
            messages,
            mode === 'compare'
              ? 'document.compare.file.aria.compare'
              : 'document.compare.file.aria.combine',
          )}
          disabled={busy}
          onFileSelect={selectFile}
        />
        <button
          type="button"
          className="work-document-compare-file"
          data-autofocus="true"
          disabled={busy}
          onClick={() => fileInputRef.current?.click()}
        >
          <FileUp size={20} aria-hidden="true" />
          <span>
            <strong>
              {file?.name ??
                officeMessage(
                  messages,
                  mode === 'compare'
                    ? 'document.compare.file.pick.compare'
                    : 'document.compare.file.pick.combine',
                )}
            </strong>
            <small>
              {file
                ? `${fileTypeLabel(messages, file.name)} · ${formatFileSize(file.size)}`
                : officeMessage(messages, 'document.compare.file.hint')}
            </small>
          </span>
        </button>

        {mode === 'compare' && (
          <label className="work-document-compare-author" htmlFor={authorId}>
            <span>
              {officeMessage(messages, 'document.compare.author.label')}
            </span>
            <OfficeTextField
              id={authorId}
              aria-label={officeMessage(
                messages,
                'document.compare.author.aria',
              )}
              value={author}
              disabled={busy}
              maxLength={256}
              onChange={(event) => {
                setAuthorTouched(true);
                setAuthor(event.currentTarget.value);
              }}
            />
            <small>
              {officeMessage(messages, 'document.compare.author.hint')}
            </small>
          </label>
        )}

        <ComparisonBoundary mode={mode} />
        {result?.status === 'unchanged' && (
          <p className="work-document-compare-status" role="status">
            {officeMessage(messages, 'document.compare.unchanged')}
          </p>
        )}
        {result?.status === 'unsupported' && (
          <div className="work-document-compare-errors" role="alert">
            <strong>
              {officeMessage(messages, 'document.compare.unsupported.title')}
            </strong>
            <ul>
              {result.diagnostics.map((diagnostic, index) => (
                <li key={`${diagnostic.code}-${index}`}>
                  {comparisonDiagnosticText(messages, diagnostic)}
                </li>
              ))}
            </ul>
          </div>
        )}
      </form>
    </Dialog>
  );
}

function ComparisonBoundary({ mode }: { mode: DocumentComparisonMode }) {
  const messages = useOfficeMessages();
  return (
    <div className="work-document-compare-boundary">
      <strong>
        {officeMessage(
          messages,
          mode === 'compare'
            ? 'document.compare.boundary.compareTitle'
            : 'document.compare.boundary.combineTitle',
        )}
      </strong>
      <p>
        {officeMessage(
          messages,
          mode === 'compare'
            ? 'document.compare.boundary.compareBody'
            : 'document.compare.boundary.combineBody',
        )}
      </p>
    </div>
  );
}

function comparisonDiagnosticText(
  messages: OfficeMessageCatalog,
  diagnostic: DocumentComparisonDiagnostic,
): string {
  const location =
    diagnostic.section === undefined
      ? ''
      : officeMessage(
          messages,
          diagnostic.block === undefined
            ? 'document.compare.diagnostic.location.section'
            : 'document.compare.diagnostic.location.sectionBlock',
          {
            section: String(diagnostic.section + 1),
            ...(diagnostic.block === undefined
              ? {}
              : { block: String(diagnostic.block + 1) }),
          },
        );
  return `${officeMessage(messages, DIAGNOSTIC_MESSAGE_KEYS[diagnostic.code])}${location}`;
}

function fileStem(name: string): string {
  return name.replace(/\.[^.]+$/, '').trim();
}

function fileTypeLabel(messages: OfficeMessageCatalog, name: string): string {
  return (
    name.split('.').at(-1)?.toLocaleUpperCase() ||
    officeMessage(messages, 'document.compare.file.fallbackType')
  );
}

function formatFileSize(bytes: number): string {
  if (bytes < 1_024) return `${bytes} B`;
  if (bytes < 1_048_576) return `${Math.max(0.1, bytes / 1_024).toFixed(1)} KB`;
  return `${Math.max(0.1, bytes / 1_048_576).toFixed(1)} MB`;
}
