import type { Editor } from '@tiptap/core';
import { MoveVertical } from 'lucide-react';
import { useState } from 'react';
import { Popover } from '../../../design-system/primitives';
import { documentParagraphSpacing } from '../work-document-paragraph-formatting';
import { OfficeNumberField } from './office-controls';

export function DocumentParagraphSpacingPopover({
  editor,
}: {
  editor: Editor;
}) {
  const spacing = documentParagraphSpacing(editor);
  const [open, setOpen] = useState(false);
  const [beforeDraft, setBeforeDraft] = useState(() =>
    pointDraft(spacing.before),
  );
  const [afterDraft, setAfterDraft] = useState(() => pointDraft(spacing.after));
  const beforeValue = pointDraft(spacing.before);
  const afterValue = pointDraft(spacing.after);
  const beforeDirty = beforeDraft !== beforeValue;
  const afterDirty = afterDraft !== afterValue;
  const customized = spacing.before !== null || spacing.after !== null;
  const commit = (key: 'before' | 'after', rawValue: string): void => {
    const value = pointValue(rawValue);
    const setDraft = key === 'before' ? setBeforeDraft : setAfterDraft;
    const current = documentParagraphSpacing(editor);
    if (rawValue.trim() && value === null) {
      setDraft(pointDraft(current[key]));
      return;
    }
    setDraft(pointDraft(value));
    if (current[key] === value) return;
    editor.commands.setDocumentParagraphSpacing(
      { ...current, [key]: value },
      { restoreFocus: false },
    );
  };
  const clear = (): void => {
    const current = documentParagraphSpacing(editor);
    setBeforeDraft('');
    setAfterDraft('');
    if (current.before === null && current.after === null) return;
    editor.commands.setDocumentParagraphSpacing(
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
      open={open}
      onOpenChange={(nextOpen) => {
        setOpen(nextOpen);
        if (!nextOpen) return;
        const current = documentParagraphSpacing(editor);
        setBeforeDraft(pointDraft(current.before));
        setAfterDraft(pointDraft(current.after));
      }}
      focusFirstOnOpen
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
            value={beforeDraft}
            min={0}
            max={720}
            step={0.5}
            placeholder="默认"
            escapeConsumer={beforeDirty}
            onValueChange={setBeforeDraft}
            onCommit={(value) => commit('before', value)}
            onCancel={
              beforeDirty ? () => setBeforeDraft(beforeValue) : undefined
            }
          />
          <span>磅</span>
        </div>
        <div className="work-document-paragraph-spacing-field">
          <span>段后</span>
          <OfficeNumberField
            ariaLabel="段后间距（磅）"
            value={afterDraft}
            min={0}
            max={720}
            step={0.5}
            placeholder="默认"
            escapeConsumer={afterDirty}
            onValueChange={setAfterDraft}
            onCommit={(value) => commit('after', value)}
            onCancel={afterDirty ? () => setAfterDraft(afterValue) : undefined}
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

function pointDraft(value: number | null): string {
  return value === null ? '' : String(value);
}

function pointValue(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(720, Math.max(0, Math.round(number * 2) / 2))
    : null;
}
