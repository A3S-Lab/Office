import { MessageSquarePlus } from 'lucide-react';
import {
  forwardRef,
  type KeyboardEvent,
  useEffect,
  useId,
  useRef,
  useState,
} from 'react';
import { Button } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { OfficeTextArea } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentCommentDraft {
  id: string;
  from: number;
  to: number;
  anchorText: string;
}

export const DocumentCommentComposer = forwardRef<
  HTMLElement,
  {
    author?: string;
    draft: DocumentCommentDraft;
    top: number;
    onCancel: () => void;
    onDirtyChange?: (dirty: boolean) => void;
    onSubmit: (text: string) => string | null;
  }
>(function DocumentCommentComposer(
  { author, draft, top, onCancel, onDirtyChange, onSubmit },
  forwardedRef,
) {
  const messages = useOfficeMessages();
  const resolvedAuthor =
    author ?? officeMessage(messages, 'document.comment.defaultAuthor');
  const titleId = useId();
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [text, setText] = useState('');
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  const dirty = Boolean(text.trim());
  useEffect(() => {
    onDirtyChange?.(dirty);
  }, [dirty, onDirtyChange]);
  useEffect(
    () => () => {
      onDirtyChange?.(false);
    },
    [onDirtyChange],
  );

  const submit = () => {
    const value = text.trim();
    if (!value) return;
    setError(onSubmit(value));
  };
  const handleKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onCancel();
      return;
    }
    if (
      event.key === 'Enter' &&
      (event.metaKey || event.ctrlKey) &&
      event.target === inputRef.current
    ) {
      event.preventDefault();
      submit();
    }
  };

  return (
    <article
      ref={forwardedRef}
      className="work-document-comment-composer"
      role="dialog"
      aria-modal="false"
      aria-labelledby={titleId}
      style={{ top: `${top}px` }}
      onKeyDown={handleKeyDown}
    >
      <header>
        <span className="work-document-comment-avatar" title={resolvedAuthor}>
          {commentAuthorInitials(
            resolvedAuthor,
            officeMessage(messages, 'document.comment.authorInitialFallback'),
          )}
        </span>
        <span className="work-document-comment-composer-heading">
          <strong id={titleId}>
            {officeMessage(messages, 'document.comment.composer.title')}
          </strong>
          <span title={draft.anchorText}>
            {resolvedAuthor} · {draft.anchorText}
          </span>
        </span>
      </header>
      <OfficeTextArea
        ref={inputRef}
        data-document-comment-body=""
        aria-label={officeMessage(
          messages,
          'document.comment.composer.bodyAria',
        )}
        value={text}
        placeholder={officeMessage(
          messages,
          'document.comment.composer.placeholder',
        )}
        onChange={(event) => {
          setText(event.target.value);
          if (error) setError(null);
        }}
      />
      {error && <p role="alert">{error}</p>}
      <footer>
        <Button size="compact" tone="quiet" onClick={onCancel}>
          {officeMessage(messages, 'document.comment.composer.cancel')}
        </Button>
        <Button
          size="compact"
          tone="primary"
          disabled={!dirty}
          onClick={submit}
        >
          <MessageSquarePlus size={13} />
          {officeMessage(messages, 'document.comment.composer.submit')}
        </Button>
      </footer>
    </article>
  );
});

function commentAuthorInitials(author: string, fallback: string): string {
  const words = author.trim().split(/\s+/).filter(Boolean);
  if (!words.length) return fallback;
  if (words.length === 1) return Array.from(words[0]).slice(0, 1).join('');
  return words
    .slice(0, 2)
    .map((word) => Array.from(word)[0])
    .join('')
    .toLocaleUpperCase();
}
