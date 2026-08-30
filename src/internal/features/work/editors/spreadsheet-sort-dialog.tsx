import { ArrowDown, ArrowUp, ListPlus, Trash2 } from 'lucide-react';
import { type FormEvent, useId, useMemo, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeCheckbox } from './office-controls';
import {
  MAX_SPREADSHEET_SORT_KEYS,
  type SpreadsheetSortDialogSource,
  type SpreadsheetSortDialogValue,
  type SpreadsheetSortKey,
} from './spreadsheet-sort';
import {
  spreadsheetSortAppearanceColumns,
  spreadsheetSortAppearanceTargets,
  spreadsheetSortAppearanceTargetsEqual,
  type SpreadsheetSortAppearanceColumn,
  type SpreadsheetSortAppearanceTarget,
} from './spreadsheet-sort-appearance';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS,
  mergeSpreadsheetSortCustomLists,
  parseSpreadsheetSortCustomList,
  spreadsheetSortCustomListsEqual,
  type SpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';
import {
  nextSpreadsheetSortKey,
  spreadsheetSortAppearanceKey,
  SpreadsheetSortOrderControls,
} from './spreadsheet-sort-order-controls';

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
  const appearanceColumns = useMemo(
    () =>
      spreadsheetSortAppearanceColumns(
        source.appearanceRows,
        source.range,
        value.hasHeader,
      ),
    [source.appearanceRows, source.range, value.hasHeader],
  );
  const [customListDraft, setCustomListDraft] =
    useState<SpreadsheetSortCustomListDraft | null>(null);
  const formId = useId();
  const nextKey =
    value.keys.length < MAX_SPREADSHEET_SORT_KEYS
      ? nextSpreadsheetSortKey(value.keys, source.columns, appearanceColumns)
      : null;

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
            disabled={!nextKey}
            onClick={() => {
              if (!nextKey) return;
              setCustomListDraft(null);
              setValue((current) => ({
                ...current,
                keys: [...current.keys, cloneSpreadsheetSortKey(nextKey)],
              }));
            }}
          >
            <ListPlus size={15} aria-hidden="true" />
            添加条件
          </Button>
          <OfficeCheckbox
            ariaLabel="数据包含标题"
            checked={value.hasHeader}
            onCheckedChange={(hasHeader) => {
              const columns = spreadsheetSortAppearanceColumns(
                source.appearanceRows,
                source.range,
                hasHeader,
              );
              setCustomListDraft(null);
              setValue((current) => ({
                ...current,
                hasHeader,
                keys: current.keys.map((key) =>
                  spreadsheetSortKeyWithColumn(
                    key,
                    key.column,
                    columns.find(
                      (candidate) => candidate.column === key.column,
                    ),
                  ),
                ),
              }));
            }}
          >
            数据包含标题
          </OfficeCheckbox>
        </div>

        <div className="work-spreadsheet-sort-levels">
          {value.keys.map((key, index) => {
            const level = index + 1;
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
                      const column = Number(event.currentTarget.value);
                      replaceKey(
                        index,
                        spreadsheetSortKeyWithColumn(
                          key,
                          column,
                          appearanceColumns.find(
                            (candidate) => candidate.column === column,
                          ),
                        ),
                      );
                    }}
                  >
                    {source.columns.map((column) => (
                      <option key={column.column} value={column.column}>
                        {column.label}
                      </option>
                    ))}
                  </select>
                </label>
                <SpreadsheetSortOrderControls
                  appearanceColumn={appearanceColumns.find(
                    (candidate) => candidate.column === key.column,
                  )}
                  customLists={customLists}
                  level={level}
                  sortKey={key}
                  onBeginCustomListEdit={(entries) =>
                    beginCustomListEdit(index, entries)
                  }
                  onChange={(replacement) => {
                    setCustomListDraft(null);
                    replaceKey(index, replacement);
                  }}
                />
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
          可按值、自定义序列、有效颜色或条件格式图标排序。值排序的空白始终置于末尾；外观排序按目标置顶/置底。新建序列仅在本次编辑器会话中复用，每次排序作为一个可撤销操作提交。
        </p>
      </form>
    </Dialog>
  );
}

function cloneSpreadsheetSortKey(key: SpreadsheetSortKey): SpreadsheetSortKey {
  if (key.sortOn === 'cell-color' || key.sortOn === 'font-color') {
    return { ...key };
  }
  if (key.sortOn === 'icon') {
    return { ...key, icon: { ...key.icon } };
  }
  if (key.customList !== undefined) {
    return { column: key.column, customList: [...key.customList] };
  }
  return { column: key.column, direction: key.direction ?? 'ascending' };
}

function spreadsheetSortKeyWithColumn(
  key: SpreadsheetSortKey,
  column: number,
  appearanceColumn: SpreadsheetSortAppearanceColumn | undefined,
): SpreadsheetSortKey {
  const target = spreadsheetSortKeyAppearanceTarget(key);
  if (target) {
    const position = key.position === 'bottom' ? 'bottom' : 'top';
    const available = spreadsheetSortAppearanceTargets(
      appearanceColumn,
      target.kind,
    );
    const selected =
      available.find((candidate) =>
        spreadsheetSortAppearanceTargetsEqual(candidate, target),
      ) ?? available[0];
    return selected
      ? spreadsheetSortAppearanceKey(column, selected, position)
      : { column, direction: 'ascending' };
  }
  return key.customList !== undefined
    ? { column, customList: [...key.customList] }
    : { column, direction: key.direction ?? 'ascending' };
}

function spreadsheetSortKeyAppearanceTarget(
  key: SpreadsheetSortKey,
): SpreadsheetSortAppearanceTarget | null {
  if (key.sortOn === 'cell-color' || key.sortOn === 'font-color') {
    return { kind: key.sortOn, color: key.color };
  }
  return key.sortOn === 'icon' ? { kind: 'icon', icon: { ...key.icon } } : null;
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
