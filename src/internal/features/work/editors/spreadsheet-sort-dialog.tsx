import { ArrowDown, ArrowUp, ListPlus, Trash2 } from 'lucide-react';
import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeCheckbox } from './office-controls';
import {
  MAX_SPREADSHEET_SORT_KEYS,
  type SpreadsheetSortDialogSource,
  type SpreadsheetSortDialogValue,
  type SpreadsheetSortKey,
} from './spreadsheet-sort';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS,
  mergeSpreadsheetSortCustomLists,
  parseSpreadsheetSortCustomList,
  spreadsheetSortCustomListsEqual,
  type SpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';

const CREATE_CUSTOM_LIST_ORDER = 'create-custom-list';
const CUSTOM_LIST_ORDER_PREFIX = 'custom-list:';

interface SpreadsheetSortCustomListDraft {
  error: string | null;
  keyIndex: number;
  text: string;
}

export function SpreadsheetSortDialog({
  source,
  restoreFocusTarget,
  onApply,
  onRememberCustomList,
  onClose,
}: {
  source: SpreadsheetSortDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetSortDialogValue) => boolean;
  onRememberCustomList?: (list: SpreadsheetSortCustomList) => void;
  onClose: () => void;
}) {
  const [value, setValue] = useState<SpreadsheetSortDialogValue>(() => ({
    hasHeader: source.value.hasHeader,
    keys: source.value.keys.map(cloneSpreadsheetSortKey),
  }));
  const [customLists, setCustomLists] = useState(() =>
    initialSpreadsheetSortCustomLists(source),
  );
  const [customListDraft, setCustomListDraft] =
    useState<SpreadsheetSortCustomListDraft | null>(null);
  const formId = useId();
  const canAdd =
    value.keys.length < source.columns.length &&
    value.keys.length < MAX_SPREADSHEET_SORT_KEYS;

  const replaceKey = (index: number, replacement: SpreadsheetSortKey) => {
    setValue((current) => ({
      ...current,
      keys: current.keys.map((key, keyIndex) =>
        keyIndex === index ? replacement : key,
      ),
    }));
  };
  const moveKey = (index: number, offset: -1 | 1) => {
    setCustomListDraft(null);
    setValue((current) => {
      const nextIndex = index + offset;
      if (nextIndex < 0 || nextIndex >= current.keys.length) return current;
      const keys = current.keys.map(cloneSpreadsheetSortKey);
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
    if (
      value.keys.length &&
      onApply({
        hasHeader: value.hasHeader,
        keys: value.keys.map(cloneSpreadsheetSortKey),
      })
    ) {
      onClose();
    }
  };

  const beginCustomListEdit = (
    keyIndex: number,
    entries?: readonly string[],
  ) => {
    setCustomListDraft({
      error: null,
      keyIndex,
      text: entries?.join('\n') ?? '',
    });
  };

  const useCustomListDraft = () => {
    if (!customListDraft) return;
    const validation = parseSpreadsheetSortCustomList(customListDraft.text);
    if (!validation.ok) {
      setCustomListDraft((current) =>
        current ? { ...current, error: validation.message } : current,
      );
      return;
    }
    const existing = customLists.find((list) =>
      spreadsheetSortCustomListsEqual(list.entries, validation.entries),
    );
    let selected = existing;
    if (!selected) {
      const sessionCount = customLists.filter(
        (list) => list.source === 'session',
      ).length;
      if (sessionCount >= MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS) {
        setCustomListDraft((current) =>
          current
            ? {
                ...current,
                error: `当前编辑器会话最多保留 ${MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS} 个自定义序列。`,
              }
            : current,
        );
        return;
      }
      const created = createSpreadsheetSortCustomList(
        validation.entries,
        'session',
      );
      if (!created) return;
      selected = created;
      setCustomLists((current) => Object.freeze([...current, created]));
      onRememberCustomList?.(created);
    }
    const key = value.keys[customListDraft.keyIndex];
    if (!key) {
      setCustomListDraft(null);
      return;
    }
    replaceKey(customListDraft.keyIndex, {
      column: key.column,
      customList: [...selected.entries],
    });
    setCustomListDraft(null);
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
              setCustomListDraft(null);
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
                    onChange={(event) => {
                      replaceKey(
                        index,
                        spreadsheetSortKeyWithColumn(
                          key,
                          Number(event.currentTarget.value),
                        ),
                      );
                    }}
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
                    value={spreadsheetSortOrderValue(key, customLists)}
                    onChange={(event) => {
                      const order = event.currentTarget.value;
                      if (order === CREATE_CUSTOM_LIST_ORDER) {
                        beginCustomListEdit(index, key.customList);
                        return;
                      }
                      if (order === 'ascending' || order === 'descending') {
                        setCustomListDraft(null);
                        replaceKey(index, {
                          column: key.column,
                          direction: order,
                        });
                        return;
                      }
                      const customListIndex = Number(
                        order.slice(CUSTOM_LIST_ORDER_PREFIX.length),
                      );
                      const customList = customLists[customListIndex];
                      if (!customList) return;
                      setCustomListDraft(null);
                      replaceKey(index, {
                        column: key.column,
                        customList: [...customList.entries],
                      });
                    }}
                  >
                    <option value="ascending">升序（A 到 Z）</option>
                    <option value="descending">降序（Z 到 A）</option>
                    <optgroup label="内置序列">
                      {customLists.map((customList, customListIndex) =>
                        customList.source === 'built-in' ? (
                          <option
                            key={`built-in:${customListIndex}`}
                            value={`${CUSTOM_LIST_ORDER_PREFIX}${customListIndex}`}
                          >
                            {customList.label}
                          </option>
                        ) : null,
                      )}
                    </optgroup>
                    {customLists.some((list) => list.source === 'session') ? (
                      <optgroup label="本次会话的序列">
                        {customLists.map((customList, customListIndex) =>
                          customList.source === 'session' ? (
                            <option
                              key={`session:${customListIndex}`}
                              value={`${CUSTOM_LIST_ORDER_PREFIX}${customListIndex}`}
                            >
                              {customList.label}
                            </option>
                          ) : null,
                        )}
                      </optgroup>
                    ) : null}
                    <option value={CREATE_CUSTOM_LIST_ORDER}>
                      新建自定义序列…
                    </option>
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
                    onClick={() => {
                      setCustomListDraft(null);
                      setValue((current) => ({
                        ...current,
                        keys: current.keys.filter(
                          (_, keyIndex) => keyIndex !== index,
                        ),
                      }));
                    }}
                  >
                    <Trash2 size={15} aria-hidden="true" />
                  </Button>
                </div>
                {customListDraft?.keyIndex === index ? (
                  <div className="work-spreadsheet-sort-custom-list-editor">
                    <label>
                      <span>自定义序列（每行一个项目）</span>
                      <textarea
                        aria-label={`排序条件 ${level} 自定义序列`}
                        aria-invalid={customListDraft.error ? true : undefined}
                        autoFocus
                        rows={5}
                        value={customListDraft.text}
                        onChange={(event) => {
                          const text = event.currentTarget.value;
                          setCustomListDraft((current) =>
                            current
                              ? {
                                  ...current,
                                  error: null,
                                  text,
                                }
                              : current,
                          );
                        }}
                      />
                    </label>
                    <p>每行、英文逗号或中文逗号可分隔一个项目。</p>
                    {customListDraft.error ? (
                      <p
                        className="work-spreadsheet-sort-custom-list-error"
                        role="alert"
                      >
                        {customListDraft.error}
                      </p>
                    ) : null}
                    <div className="work-spreadsheet-sort-custom-list-actions">
                      <Button
                        tone="primary"
                        type="button"
                        onClick={useCustomListDraft}
                      >
                        使用序列
                      </Button>
                      <Button
                        tone="quiet"
                        type="button"
                        onClick={() => setCustomListDraft(null)}
                      >
                        取消编辑
                      </Button>
                    </div>
                  </div>
                ) : key.customList !== undefined ? (
                  <div className="work-spreadsheet-sort-custom-list-preview">
                    <code>{key.customList.join(' → ')}</code>
                    <Button
                      tone="quiet"
                      type="button"
                      onClick={() => beginCustomListEdit(index, key.customList)}
                    >
                      编辑序列
                    </Button>
                  </div>
                ) : null}
              </fieldset>
            );
          })}
        </div>
        <p className="work-spreadsheet-sort-note">
          按单元格值或自定义序列排序；空白单元格始终置于末尾。新建序列仅在本次编辑器会话中复用，每次排序作为一个可撤销操作提交。
        </p>
      </form>
    </Dialog>
  );
}

function cloneSpreadsheetSortKey(key: SpreadsheetSortKey): SpreadsheetSortKey {
  return key.customList !== undefined
    ? { column: key.column, customList: [...key.customList] }
    : { column: key.column, direction: key.direction };
}

function spreadsheetSortKeyWithColumn(
  key: SpreadsheetSortKey,
  column: number,
): SpreadsheetSortKey {
  return key.customList !== undefined
    ? { column, customList: [...key.customList] }
    : { column, direction: key.direction };
}

function spreadsheetSortOrderValue(
  key: SpreadsheetSortKey,
  customLists: readonly SpreadsheetSortCustomList[],
): string {
  if (key.customList === undefined) return key.direction;
  const index = customLists.findIndex((customList) =>
    spreadsheetSortCustomListsEqual(customList.entries, key.customList ?? []),
  );
  return index < 0
    ? CREATE_CUSTOM_LIST_ORDER
    : `${CUSTOM_LIST_ORDER_PREFIX}${index}`;
}

function initialSpreadsheetSortCustomLists(
  source: SpreadsheetSortDialogSource,
): readonly SpreadsheetSortCustomList[] {
  const candidates: SpreadsheetSortCustomList[] = [];
  for (const key of source.value.keys) {
    if (key.customList === undefined) continue;
    const list = createSpreadsheetSortCustomList(key.customList, 'session');
    if (list) candidates.push(list);
  }
  candidates.push(...source.customLists);
  return mergeSpreadsheetSortCustomLists(candidates);
}
