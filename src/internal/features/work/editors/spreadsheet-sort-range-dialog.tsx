import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import type {
  SpreadsheetSortRangeChoice,
  SpreadsheetSortRangeDialogSource,
} from './spreadsheet-sort';

export function SpreadsheetSortRangeDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  source: SpreadsheetSortRangeDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (choice: SpreadsheetSortRangeChoice) => boolean;
  onClose: () => void;
}) {
  const [choice, setChoice] = useState<SpreadsheetSortRangeChoice>(() =>
    source.canSortExpandedRange ? 'expand' : 'selection',
  );
  const formId = useId();
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (onApply(choice)) onClose();
  };

  return (
    <Dialog
      title="排序提醒"
      description={`${source.sheetName} 中选定区域旁边还有数据。`}
      className="work-spreadsheet-sort-range-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button tone="primary" type="submit" form={formId}>
            排序
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-spreadsheet-sort-range-options">
          <legend>请选择排序范围</legend>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-range"
              value="expand"
              checked={choice === 'expand'}
              disabled={!source.canSortExpandedRange}
              onChange={() => setChoice('expand')}
            />
            <span>
              <strong>扩展选定区域</strong>
              <small>让相邻列随整行一起移动，避免数据关系错位。</small>
              <code>{source.expandedRangeReference}</code>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-range"
              value="selection"
              checked={choice === 'selection'}
              disabled={!source.canSortSelection}
              onChange={() => setChoice('selection')}
            />
            <span>
              <strong>以当前选定区域排序</strong>
              <small>只移动当前矩形中的单元格。</small>
              <code>{source.selectedRangeReference}</code>
            </span>
          </label>
        </fieldset>
      </form>
    </Dialog>
  );
}
