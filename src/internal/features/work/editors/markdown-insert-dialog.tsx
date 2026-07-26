import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeTextField } from './office-controls';

export type MarkdownInsertDialogRequest =
  | {
      kind: 'link';
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
  const formId = useId();
  const textFieldId = useId();
  const sourceFieldId = useId();
  const [source, setSource] = useState(request.source);
  const [text, setText] = useState(
    request.kind === 'link' ? request.label : request.altText,
  );
  const sourceReady = isCompletedSource(source);
  const submit = (event: FormEvent) => {
    event.preventDefault();
    const normalizedSource = source.trim();
    if (!isCompletedSource(normalizedSource)) return;
    onSubmit(
      request.kind === 'link'
        ? {
            kind: 'link',
            label: text.trim() || normalizedSource,
            source: normalizedSource,
          }
        : {
            kind: 'image',
            altText: text.trim(),
            source: normalizedSource,
          },
    );
  };
  const title = request.kind === 'link' ? '添加链接' : '插入图片';

  return (
    <Dialog
      title={title}
      className="work-office-dialog work-markdown-insert-dialog"
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
            disabled={!sourceReady}
          >
            {title}
          </Button>
        </>
      }
    >
      <form id={formId} className="work-markdown-insert-form" onSubmit={submit}>
        <label className="work-office-dialog-field" htmlFor={textFieldId}>
          <span className="work-office-dialog-field-label">
            {request.kind === 'link' ? '显示文字' : '替代文字（可选）'}
          </span>
          <OfficeTextField
            id={textFieldId}
            aria-label={
              request.kind === 'link' ? '显示文字' : '替代文字（可选）'
            }
            value={text}
            onChange={(event) => setText(event.target.value)}
          />
        </label>
        <label className="work-office-dialog-field" htmlFor={sourceFieldId}>
          <span className="work-office-dialog-field-label">
            {request.kind === 'link' ? '链接地址' : '图片地址'}
          </span>
          <OfficeTextField
            id={sourceFieldId}
            aria-label={request.kind === 'link' ? '链接地址' : '图片地址'}
            inputMode="url"
            placeholder="https://"
            value={source}
            onChange={(event) => setSource(event.target.value)}
          />
        </label>
      </form>
    </Dialog>
  );
}

function isCompletedSource(source: string): boolean {
  const value = source.trim();
  return Boolean(value && value !== 'https://');
}
