import type { Editor } from '@tiptap/core';
import { MoveVertical } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import {
  documentParagraphSpacing,
  setDocumentParagraphSpacing,
} from '../work-document-paragraph-formatting';
import { OfficeNumberField } from './office-controls';

export function DocumentParagraphSpacingPopover({
  editor,
}: {
  editor: Editor;
}) {
  const spacing = documentParagraphSpacing(editor);
  const customized = spacing.before !== null || spacing.after !== null;
  const update = (key: 'before' | 'after', rawValue: string): void => {
    const value = pointValue(rawValue);
    if (rawValue.trim() && value === null) return;
    const current = documentParagraphSpacing(editor);
    setDocumentParagraphSpacing(
      editor,
      { ...current, [key]: value },
      { restoreFocus: false },
    );
  };
  const clear = (): void => {
    const current = documentParagraphSpacing(editor);
    setDocumentParagraphSpacing(
      editor,
      { ...current, before: null, after: null },
      { restoreFocus: false },
    );
  };

  return (
    <Popover
      label="段落间距"
      panelLabel="段落间距选项"
      panelRole="dialog"
      portal
      className="work-document-paragraph-spacing-popover"
      panelClassName="work-document-paragraph-spacing-panel"
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label${customized || open ? ' active' : ''}`}
          aria-pressed={customized}
        >
          <MoveVertical size={19} />
          <span>段落间距</span>
        </button>
      )}
    >
      <fieldset>
        <legend>段落间距</legend>
        <div className="work-document-paragraph-spacing-field">
          <span>段前</span>
          <OfficeNumberField
            ariaLabel="段前间距（磅）"
            value={spacing.before ?? ''}
            min={0}
            max={720}
            step={0.5}
            placeholder="默认"
            onValueChange={(value) => update('before', value)}
          />
          <span>磅</span>
        </div>
        <div className="work-document-paragraph-spacing-field">
          <span>段后</span>
          <OfficeNumberField
            ariaLabel="段后间距（磅）"
            value={spacing.after ?? ''}
            min={0}
            max={720}
            step={0.5}
            placeholder="默认"
            onValueChange={(value) => update('after', value)}
          />
          <span>磅</span>
        </div>
        <button
          type="button"
          className="work-document-paragraph-spacing-reset"
          aria-label="恢复默认间距"
          onClick={clear}
        >
          恢复默认
        </button>
      </fieldset>
    </Popover>
  );
}

function pointValue(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(720, Math.max(0, Math.round(number * 2) / 2))
    : null;
}
