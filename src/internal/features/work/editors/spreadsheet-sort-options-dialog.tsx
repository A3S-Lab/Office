import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import type { SpreadsheetSortOrientation } from './spreadsheet-sort';

export function SpreadsheetSortOptionsDialog({
  orientation,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  orientation: SpreadsheetSortOrientation;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (orientation: SpreadsheetSortOrientation) => void;
  onClose: () => void;
}) {
  const [draft, setDraft] = useState(orientation);
  const formId = useId();
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    onApply(draft);
  };

  return (
    <Dialog
      title="排序选项"
      description="选择数据在选定区域内的排序方向。"
      className="work-spreadsheet-sort-options-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button tone="primary" type="submit" form={formId}>
            确定
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-spreadsheet-sort-options">
          <legend>方向</legend>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-orientation"
              value="top-to-bottom"
              checked={draft === 'top-to-bottom'}
              data-autofocus={draft === 'top-to-bottom' ? '' : undefined}
              onChange={() => setDraft('top-to-bottom')}
            />
            <span>
              <strong>按列排序</strong>
              <small>根据所选列，从上到下移动整行。</small>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-orientation"
              value="left-to-right"
              checked={draft === 'left-to-right'}
              data-autofocus={draft === 'left-to-right' ? '' : undefined}
              onChange={() => setDraft('left-to-right')}
            />
            <span>
              <strong>按行排序</strong>
              <small>根据所选行，从左到右移动整列。</small>
            </span>
          </label>
        </fieldset>
      </form>
    </Dialog>
  );
}
