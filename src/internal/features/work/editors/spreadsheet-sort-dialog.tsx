import { ArrowDown, ArrowUp, ListPlus, Settings2, Trash2 } from 'lucide-react';
import { type FormEvent, useId, useMemo, useRef, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeCheckbox } from './office-controls';
import {
  MAX_SPREADSHEET_SORT_KEYS,
  type SpreadsheetSortDialogSource,
  type SpreadsheetSortDialogValue,
  type SpreadsheetSortKey,
} from './spreadsheet-sort';
import {
  spreadsheetSortAppearanceFields,
  spreadsheetSortAppearanceTargets,
  spreadsheetSortAppearanceTargetsEqual,
  type SpreadsheetSortAppearanceField,
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
import { SpreadsheetSortCustomListEditor } from './spreadsheet-sort-custom-list-editor';
import { SpreadsheetSortOptionsDialog } from './spreadsheet-sort-options-dialog';
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
    orientation: source.value.orientation,
  }));
  const [verticalHasHeader, setVerticalHasHeader] = useState(
    source.value.hasHeader,
  );
  const [optionsOpen, setOptionsOpen] = useState(false);
  const [customLists, setCustomLists] = useState(() =>
    initialSpreadsheetSortCustomLists(source),
  );
  const fields =
    value.orientation === 'top-to-bottom' ? source.columns : source.rows;
  const appearanceFields = useMemo(
    () =>
      spreadsheetSortAppearanceFields(
        source.appearanceRows,
        source.range,
        value.orientation,
        value.hasHeader,
      ),
    [source.appearanceRows, source.range, value.hasHeader, value.orientation],
  );
  const [customListDraft, setCustomListDraft] =
    useState<SpreadsheetSortCustomListDraft | null>(null);
  const formId = useId();
  const optionsButtonRef = useRef<HTMLButtonElement>(null);
  const nextKey =
    value.keys.length < MAX_SPREADSHEET_SORT_KEYS
      ? nextSpreadsheetSortKey(value.keys, fields, appearanceFields)
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
        orientation: value.orientation,
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
      index: key.index,
      customList: [...selected.entries],
    });
    setCustomListDraft(null);
  };

  const applyOrientation = (
    orientation: SpreadsheetSortDialogValue['orientation'],
  ) => {
    setOptionsOpen(false);
    if (orientation === value.orientation) return;
    const nextFields =
      orientation === 'top-to-bottom' ? source.columns : source.rows;
    const preferredIndex =
      orientation === 'top-to-bottom'
        ? (source.value.keys[0]?.index ?? nextFields[0]?.index)
        : source.activeRow;
    const index =
      nextFields.find((field) => field.index === preferredIndex)?.index ??
      nextFields[0]?.index;
    if (index === undefined) return;
    setCustomListDraft(null);
    setValue({
      orientation,
      hasHeader: orientation === 'top-to-bottom' ? verticalHasHeader : false,
      keys: [{ index, direction: 'ascending' }],
    });
  };

  return (
    <>
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
            <Button
              ref={optionsButtonRef}
              tone="quiet"
              type="button"
              onClick={() => setOptionsOpen(true)}
            >
              <Settings2 size={15} aria-hidden="true" />
              选项…
            </Button>
            <OfficeCheckbox
              ariaLabel="数据包含标题"
              checked={value.hasHeader}
              disabled={value.orientation === 'left-to-right'}
              onCheckedChange={(hasHeader) => {
                const nextAppearanceFields = spreadsheetSortAppearanceFields(
                  source.appearanceRows,
                  source.range,
                  'top-to-bottom',
                  hasHeader,
                );
                setVerticalHasHeader(hasHeader);
                setCustomListDraft(null);
                setValue((current) => ({
                  ...current,
                  hasHeader,
                  keys: current.keys.map((key) =>
                    spreadsheetSortKeyWithIndex(
                      key,
                      key.index,
                      nextAppearanceFields.find(
                        (candidate) => candidate.index === key.index,
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
                  key={`${index}:${key.index}`}
                >
                  <legend>
                    {index === 0 ? '主要关键字' : `次要关键字 ${index}`}
                  </legend>
                  <label>
                    <span>
                      {value.orientation === 'top-to-bottom' ? '列' : '行'}
                    </span>
                    <select
                      aria-label={`排序条件 ${level} ${
                        value.orientation === 'top-to-bottom' ? '列' : '行'
                      }`}
                      value={key.index}
                      onChange={(event) => {
                        const fieldIndex = Number(event.currentTarget.value);
                        replaceKey(
                          index,
                          spreadsheetSortKeyWithIndex(
                            key,
                            fieldIndex,
                            appearanceFields.find(
                              (candidate) => candidate.index === fieldIndex,
                            ),
                          ),
                        );
                      }}
                    >
                      {fields.map((field) => (
                        <option key={field.index} value={field.index}>
                          {field.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <SpreadsheetSortOrderControls
                    appearanceField={appearanceFields.find(
                      (candidate) => candidate.index === key.index,
                    )}
                    customLists={customLists}
                    level={level}
                    orientation={value.orientation}
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
                    <SpreadsheetSortCustomListEditor
                      error={customListDraft.error}
                      level={level}
                      text={customListDraft.text}
                      onCancel={() => setCustomListDraft(null)}
                      onChange={(text) =>
                        setCustomListDraft((current) =>
                          current ? { ...current, error: null, text } : current,
                        )
                      }
                      onUse={useCustomListDraft}
                    />
                  ) : key.customList !== undefined ? (
                    <div className="work-spreadsheet-sort-custom-list-preview">
                      <code>{key.customList.join(' → ')}</code>
                      <Button
                        tone="quiet"
                        type="button"
                        onClick={() =>
                          beginCustomListEdit(index, key.customList)
                        }
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
            可按值、自定义序列、有效颜色或条件格式图标排序。按列排序移动整行；按行排序移动整列且不保留标题列。值排序的空白始终置于末尾，新建序列仅在本次编辑器会话中复用，每次排序作为一个可撤销操作提交。
          </p>
        </form>
      </Dialog>
      {optionsOpen ? (
        <SpreadsheetSortOptionsDialog
          orientation={value.orientation}
          restoreFocusTarget={() => optionsButtonRef.current}
          onApply={applyOrientation}
          onClose={() => setOptionsOpen(false)}
        />
      ) : null}
    </>
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
    return { index: key.index, customList: [...key.customList] };
  }
  return { index: key.index, direction: key.direction ?? 'ascending' };
}

function spreadsheetSortKeyWithIndex(
  key: SpreadsheetSortKey,
  index: number,
  appearanceField: SpreadsheetSortAppearanceField | undefined,
): SpreadsheetSortKey {
  const target = spreadsheetSortKeyAppearanceTarget(key);
  if (target) {
    const position = key.position === 'last' ? 'last' : 'first';
    const available = spreadsheetSortAppearanceTargets(
      appearanceField,
      target.kind,
    );
    const selected =
      available.find((candidate) =>
        spreadsheetSortAppearanceTargetsEqual(candidate, target),
      ) ?? available[0];
    return selected
      ? spreadsheetSortAppearanceKey(index, selected, position)
      : { index, direction: 'ascending' };
  }
  return key.customList !== undefined
    ? { index, customList: [...key.customList] }
    : { index, direction: key.direction ?? 'ascending' };
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
