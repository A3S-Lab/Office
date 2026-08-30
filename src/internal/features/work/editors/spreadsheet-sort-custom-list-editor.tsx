import { useEffect, useRef } from 'react';
import { Button } from '../../../design-system/primitives';

export function SpreadsheetSortCustomListEditor({
  error,
  level,
  text,
  onCancel,
  onChange,
  onUse,
}: {
  error: string | null;
  level: number;
  text: string;
  onCancel: () => void;
  onChange: (text: string) => void;
  onUse: () => void;
}) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  useEffect(() => textareaRef.current?.focus(), []);

  return (
    <div className="work-spreadsheet-sort-custom-list-editor">
      <label>
        <span>自定义序列（每行一个项目）</span>
        <textarea
          ref={textareaRef}
          aria-label={`排序条件 ${level} 自定义序列`}
          aria-invalid={error ? true : undefined}
          rows={5}
          value={text}
          onChange={(event) => onChange(event.currentTarget.value)}
        />
      </label>
      <p>每行、英文逗号或中文逗号可分隔一个项目。</p>
      {error ? (
        <p className="work-spreadsheet-sort-custom-list-error" role="alert">
          {error}
        </p>
      ) : null}
      <div className="work-spreadsheet-sort-custom-list-actions">
        <Button tone="primary" type="button" onClick={onUse}>
          使用序列
        </Button>
        <Button tone="quiet" type="button" onClick={onCancel}>
          取消编辑
        </Button>
      </div>
    </div>
  );
}
