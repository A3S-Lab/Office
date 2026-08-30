import { ArrowDown, ArrowUp, ListPlus, Trash2 } from 'lucide-react';
import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeCheckbox } from './office-controls';
import {
  MAX_SPREADSHEET_SORT_KEYS,
  type SpreadsheetSortDialogSource,
  type SpreadsheetSortDialogValue,
  type SpreadsheetSortDirection,
  type SpreadsheetSortKey,
} from './spreadsheet-sort';

export function SpreadsheetSortDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  source: SpreadsheetSortDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetSortDialogValue) => boolean;
  onClose: () => void;
}) {
  const [value, setValue] = useState<SpreadsheetSortDialogValue>(() => ({
    hasHeader: source.value.hasHeader,
    keys: source.value.keys.map((key) => ({ ...key })),
  }));
  const formId = useId();
  const canAdd =
    value.keys.length < source.columns.length &&
    value.keys.length < MAX_SPREADSHEET_SORT_KEYS;

  const updateKey = (index: number, patch: Partial<SpreadsheetSortKey>) => {
    setValue((current) => ({
      ...current,
      keys: current.keys.map((key, keyIndex) =>
        keyIndex === index ? { ...key, ...patch } : key,
      ),
    }));
  };
  const moveKey = (index: number, offset: -1 | 1) => {
    setValue((current) => {
      const nextIndex = index + offset;
      if (nextIndex < 0 || nextIndex >= current.keys.length) return current;
      const keys = current.keys.map((key) => ({ ...key }));
      const currentKey = keys[index];
      const nextKey = keys[nextIndex];
      if (!currentKey || !nextKey) return current;
      keys[index] = nextKey;
      keys[nextIndex] = currentKey;
      return { ...current, keys };
    });
  };
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (value.keys.length && onApply(value)) onClose();
  };

  return (
    <Dialog
      title="自定义排序"
      description={`${source.sheetName}!${source.rangeReference}`}
      className="work-spreadsheet-sort-dialog"
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
        <div className="work-spreadsheet-sort-toolbar">
          <Button
            tone="quiet"
            type="button"
            disabled={!canAdd}
            onClick={() => {
              const used = new Set(value.keys.map((key) => key.column));
              const next = source.columns.find(
                (column) => !used.has(column.column),
              );
              if (!next) return;
              setValue((current) => ({
                ...current,
                keys: [
                  ...current.keys,
                  { column: next.column, direction: 'ascending' },
                ],
              }));
            }}
          >
            <ListPlus size={15} aria-hidden="true" />
            添加条件
          </Button>
          <OfficeCheckbox
            ariaLabel="数据包含标题"
            checked={value.hasHeader}
            onCheckedChange={(hasHeader) =>
              setValue((current) => ({ ...current, hasHeader }))
            }
          >
            数据包含标题
          </OfficeCheckbox>
        </div>

        <div className="work-spreadsheet-sort-levels">
          {value.keys.map((key, index) => {
            const level = index + 1;
            const usedByOthers = new Set(
              value.keys
                .filter((_, keyIndex) => keyIndex !== index)
                .map((candidate) => candidate.column),
            );
            return (
              <fieldset
                className="work-spreadsheet-sort-level"
                key={`${index}:${key.column}`}
              >
                <legend>
                  {index === 0 ? '主要关键字' : `次要关键字 ${index}`}
                </legend>
                <label>
                  <span>列</span>
                  <select
                    aria-label={`排序条件 ${level} 列`}
                    value={key.column}
                    onChange={(event) =>
                      updateKey(index, {
                        column: Number(event.currentTarget.value),
                      })
                    }
                  >
                    {source.columns.map((column) => (
                      <option
                        key={column.column}
                        value={column.column}
                        disabled={usedByOthers.has(column.column)}
                      >
                        {column.label}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  <span>次序</span>
                  <select
                    aria-label={`排序条件 ${level} 次序`}
                    value={key.direction}
                    onChange={(event) =>
                      updateKey(index, {
                        direction: event.currentTarget
                          .value as SpreadsheetSortDirection,
                      })
                    }
                  >
                    <option value="ascending">升序（A 到 Z）</option>
                    <option value="descending">降序（Z 到 A）</option>
                  </select>
                </label>
                <div className="work-spreadsheet-sort-level-actions">
                  <Button
                    tone="quiet"
                    type="button"
                    aria-label={`上移条件 ${level}`}
                    title="提高排序优先级"
                    disabled={index === 0}
                    onClick={() => moveKey(index, -1)}
                  >
                    <ArrowUp size={15} aria-hidden="true" />
                  </Button>
                  <Button
                    tone="quiet"
                    type="button"
                    aria-label={`下移条件 ${level}`}
                    title="降低排序优先级"
                    disabled={index === value.keys.length - 1}
                    onClick={() => moveKey(index, 1)}
                  >
                    <ArrowDown size={15} aria-hidden="true" />
                  </Button>
                  <Button
                    tone="quiet"
                    type="button"
                    aria-label={`删除条件 ${level}`}
                    title="删除排序条件"
                    disabled={value.keys.length === 1}
                    onClick={() =>
                      setValue((current) => ({
                        ...current,
                        keys: current.keys.filter(
                          (_, keyIndex) => keyIndex !== index,
                        ),
                      }))
                    }
                  >
                    <Trash2 size={15} aria-hidden="true" />
                  </Button>
                </div>
              </fieldset>
            );
          })}
        </div>
        <p className="work-spreadsheet-sort-note">
          按单元格值排序；空白单元格始终置于末尾。每次排序作为一个可撤销操作提交。
        </p>
      </form>
    </Dialog>
  );
}
