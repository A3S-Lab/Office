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
import { OfficeTextArea } from './office-controls';

export interface DocumentCommentDraft {
  id: string;
  from: number;
  to: number;
  anchorText: string;
}

export const DocumentCommentComposer = forwardRef<
  HTMLElement,
  {
    draft: DocumentCommentDraft;
    top: number;
    onCancel: () => void;
    onDirtyChange?: (dirty: boolean) => void;
    onSubmit: (text: string) => string | null;
  }
>(function DocumentCommentComposer(
  { draft, top, onCancel, onDirtyChange, onSubmit },
  forwardedRef,
) {
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
        <span className="work-document-comment-avatar">我</span>
        <span className="work-document-comment-composer-heading">
          <strong id={titleId}>添加批注</strong>
          <span title={draft.anchorText}>{draft.anchorText}</span>
        </span>
      </header>
      <OfficeTextArea
        ref={inputRef}
        aria-label="批注内容"
        value={text}
        placeholder="输入批注…"
        onChange={(event) => {
          setText(event.target.value);
          if (error) setError(null);
        }}
      />
      {error && <p role="alert">{error}</p>}
      <footer>
        <Button size="compact" tone="quiet" onClick={onCancel}>
          取消
        </Button>
        <Button
          size="compact"
          tone="primary"
          disabled={!dirty}
          onClick={submit}
        >
          <MessageSquarePlus size={13} />
          添加批注
        </Button>
      </footer>
    </article>
  );
});
