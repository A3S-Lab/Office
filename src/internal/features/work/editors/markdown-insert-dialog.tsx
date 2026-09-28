import { type FormEvent, useId, useRef, useState } from 'react';
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
import { Button, Dialog } from '../../../design-system/primitives';
import {
  DOCUMENT_LINK_VALIDATION_MESSAGE,
  normalizeDocumentHref,
} from '../work-document-links';
import { OfficeTextField } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export function markdownImageValidationMessage(
  messages = resolveOfficeMessages(),
): string {
  return officeMessage(messages, 'markdown.insert.imageUrlError');
}

/** @deprecated Prefer markdownImageValidationMessage(messages). */
export const MARKDOWN_IMAGE_VALIDATION_MESSAGE =
  markdownImageValidationMessage();

export type MarkdownInsertDialogRequest =
  | {
      kind: 'link';
      action: 'edit' | 'insert';
      label: string;
      source: string;
    }
  | {
      kind: 'image';
      altText: string;
      source: string;
    };

export type MarkdownInsertDialogResult = MarkdownInsertDialogRequest;

export function MarkdownInsertDialog({
  request,
  restoreFocusTarget,
  onClose,
  onSubmit,
}: {
  request: MarkdownInsertDialogRequest;
  restoreFocusTarget: () => HTMLElement | null;
  onClose: () => void;
  onSubmit: (result: MarkdownInsertDialogResult) => void;
}) {
  const messages = useOfficeMessages();
  const formId = useId();
  const textFieldId = useId();
  const sourceFieldId = useId();
  const submittedRef = useRef(false);
  const [source, setSource] = useState(request.source);
  const [text, setText] = useState(
    request.kind === 'link' ? request.label : request.altText,
  );
  const normalizedSource = normalizeMarkdownInsertSource(request.kind, source);
  const sourceReady = normalizedSource !== null;
  const sourceError = markdownInsertSourceError(request.kind, source, messages);
  const sourceErrorId = `${sourceFieldId}-error`;
  const submit = (event: FormEvent) => {
    event.preventDefault();
    if (submittedRef.current) return;
    const completedSource = normalizeMarkdownInsertSource(request.kind, source);
    if (!completedSource) return;
    submittedRef.current = true;
    onSubmit(
      request.kind === 'link'
        ? {
            kind: 'link',
            action: request.action,
            label: text.trim() || completedSource,
            source: completedSource,
          }
        : {
            kind: 'image',
            altText: text.trim(),
            source: completedSource,
          },
    );
  };
  const title =
    request.kind === 'link'
      ? request.action === 'edit'
        ? officeMessage(messages, 'markdown.insert.editLink')
        : officeMessage(messages, 'markdown.insert.addLink')
      : officeMessage(messages, 'markdown.insert.insertImage');
  const submitLabel =
    request.kind === 'link'
      ? request.action === 'edit'
        ? officeMessage(messages, 'markdown.insert.save')
        : officeMessage(messages, 'markdown.insert.add')
      : officeMessage(messages, 'markdown.insert.insert');
  const focusText = request.kind === 'link' && !request.label.trim();

  return (
    <Dialog
      title={title}
      className="work-office-dialog work-markdown-insert-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'markdown.insert.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!sourceReady}
          >
            {submitLabel}
          </Button>
        </>
      }
    >
      <form id={formId} className="work-markdown-insert-form" onSubmit={submit}>
        <label className="work-office-dialog-field" htmlFor={textFieldId}>
          <span className="work-office-dialog-field-label">
            {request.kind === 'link' ? officeMessage(messages, 'markdown.insert.linkText') : officeMessage(messages, 'markdown.insert.altText')}
          </span>
          <OfficeTextField
            id={textFieldId}
            aria-label={
              request.kind === 'link' ? officeMessage(messages, 'markdown.insert.linkText') : officeMessage(messages, 'markdown.insert.altText')
            }
            data-autofocus={focusText || undefined}
            value={text}
            onChange={(event) => setText(event.target.value)}
          />
        </label>
        <label className="work-office-dialog-field" htmlFor={sourceFieldId}>
          <span className="work-office-dialog-field-label">
            {request.kind === 'link' ? officeMessage(messages, 'markdown.insert.linkUrl') : officeMessage(messages, 'markdown.insert.imageUrl')}
          </span>
          <OfficeTextField
            id={sourceFieldId}
            aria-label={request.kind === 'link' ? officeMessage(messages, 'markdown.insert.linkUrl') : officeMessage(messages, 'markdown.insert.imageUrl')}
            aria-describedby={sourceError ? sourceErrorId : undefined}
            aria-invalid={sourceError ? true : undefined}
            inputMode="url"
            placeholder="https://"
            data-autofocus={!focusText || undefined}
            value={source}
            onChange={(event) => setSource(event.target.value)}
          />
          {sourceError && (
            <span
              id={sourceErrorId}
              className="work-office-dialog-field-error"
              role="alert"
            >
              {sourceError}
            </span>
          )}
        </label>
      </form>
    </Dialog>
  );
}

export function normalizeMarkdownInsertSource(
  kind: MarkdownInsertDialogRequest['kind'],
  source: string,
): string | null {
  const value = source.trim();
  if (!value || /[\u0000-\u001f\u007f\\<>]/u.test(value)) return null;
  if (kind === 'link') {
    const normalized = normalizeDocumentHref(value);
    return normalized ? normalizeMarkdownSourceSpaces(normalized) : null;
  }

  if (/^https?:/iu.test(value)) {
    const normalized = normalizeDocumentHref(value);
    return normalized && /^https?:/iu.test(normalized)
      ? normalizeMarkdownSourceSpaces(normalized)
      : null;
  }

  try {
    const base = new URL('https://a3s-office.invalid/');
    const resolved = new URL(value, base);
    const path = value.split(/[?#]/u, 1)[0] ?? '';
    return resolved.protocol === base.protocol &&
      resolved.origin === base.origin &&
      /[^\s./]/u.test(path)
      ? normalizeMarkdownSourceSpaces(value)
      : null;
  } catch {
    return null;
  }
}

function normalizeMarkdownSourceSpaces(source: string): string {
  return source.replace(/ /gu, '%20');
}

function markdownInsertSourceError(
  kind: MarkdownInsertDialogRequest['kind'],
  source: string,
  messages = resolveOfficeMessages(),
): string | null {
  const value = source.trim();
  if (!value || value === 'https://') return null;
  if (normalizeMarkdownInsertSource(kind, value)) return null;
  return kind === 'link'
    ? DOCUMENT_LINK_VALIDATION_MESSAGE
    : markdownImageValidationMessage(messages);
}
