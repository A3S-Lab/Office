import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import type { SpreadsheetSortOptions } from './spreadsheet-sort';

export function SpreadsheetSortOptionsDialog({
  value,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  value: SpreadsheetSortOptions;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetSortOptions) => void;
  onClose: () => void;
}) {
  const [draft, setDraft] = useState(value);
  const formId = useId();
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    onApply(draft);
  };

  return (
    <Dialog
      title="排序选项"
      description="选择文本比较规则和数据在选定区域内的排序方向。"
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
      <form
        id={formId}
        className="work-spreadsheet-sort-options-form"
        onSubmit={submit}
      >
        <fieldset className="work-spreadsheet-sort-options">
          <legend>文本比较</legend>
          <label>
            <input
              type="checkbox"
              aria-label="区分大小写"
              checked={draft.caseSensitive}
              onChange={(event) => {
                const caseSensitive = event.currentTarget.checked;
                setDraft((current) => ({
                  ...current,
                  caseSensitive,
                }));
              }}
            />
            <span>
              <strong>区分大小写</strong>
              <small>升序时，同一字母的小写形式排在大写形式之前。</small>
            </span>
          </label>
        </fieldset>
        <fieldset className="work-spreadsheet-sort-options">
          <legend>方法</legend>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-text-method"
              value="pinyin"
              aria-label="拼音排序"
              checked={draft.textMethod === 'pinyin'}
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  textMethod: 'pinyin',
                }))
              }
            />
            <span>
              <strong>拼音排序</strong>
              <small>按汉字拼音和文本字符顺序比较。</small>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-text-method"
              value="stroke"
              aria-label="笔画排序"
              checked={draft.textMethod === 'stroke'}
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  textMethod: 'stroke',
                }))
              }
            />
            <span>
              <strong>笔画排序</strong>
              <small>按汉字笔画顺序比较，适合中文姓名和目录。</small>
            </span>
          </label>
        </fieldset>
        <fieldset className="work-spreadsheet-sort-options">
          <legend>方向</legend>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-orientation"
              value="top-to-bottom"
              checked={draft.orientation === 'top-to-bottom'}
              data-autofocus={
                draft.orientation === 'top-to-bottom' ? '' : undefined
              }
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  orientation: 'top-to-bottom',
                }))
              }
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
              checked={draft.orientation === 'left-to-right'}
              data-autofocus={
                draft.orientation === 'left-to-right' ? '' : undefined
              }
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  orientation: 'left-to-right',
                }))
              }
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
