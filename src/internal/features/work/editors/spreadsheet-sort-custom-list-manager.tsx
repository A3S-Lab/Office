import { ArrowDown, ArrowUp, ListPlus, Save, Trash2 } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS,
  parseSpreadsheetSortCustomList,
  spreadsheetSortCustomListsEqual,
  type SpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';

interface ManagedSpreadsheetSortCustomList {
  entries: readonly string[];
  id: number;
  originalEntries: readonly string[] | null;
}

export interface SpreadsheetSortCustomListManagementChange {
  next: readonly string[] | null;
  previous: readonly string[];
}

export interface SpreadsheetSortCustomListManagementResult {
  changes: readonly SpreadsheetSortCustomListManagementChange[];
  lists: readonly (readonly string[])[];
}

export function SpreadsheetSortCustomListManagerDialog({
  customLists,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  customLists: readonly SpreadsheetSortCustomList[];
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetSortCustomListManagementResult) => void;
  onClose: () => void;
}) {
  const builtInLists = useMemo(
    () => customLists.filter((list) => list.source === 'built-in'),
    [customLists],
  );
  const initialUserListsRef = useRef<
    readonly ManagedSpreadsheetSortCustomList[] | null
  >(null);
  initialUserListsRef.current ??= customLists
    .filter((list) => list.source !== 'built-in')
    .map((list, index) => ({
      id: index,
      entries: Object.freeze([...list.entries]),
      originalEntries: Object.freeze([...list.entries]),
    }));
  const initialUserLists = initialUserListsRef.current;
  const nextIdRef = useRef(initialUserLists.length);
  const listRef = useRef<HTMLSelectElement>(null);
  const [userLists, setUserLists] = useState<
    readonly ManagedSpreadsheetSortCustomList[]
  >(() => initialUserLists.map(cloneManagedCustomList));
  const [selection, setSelection] = useState('built-in:0');
  const [text, setText] = useState(builtInLists[0]?.entries.join('\n') ?? '');
  const [error, setError] = useState<string | null>(null);
  const userListLabels = useMemo(
    () =>
      new Map(
        userLists.map((list) => [
          list.id,
          managedCustomListLabel(list.entries),
        ]),
      ),
    [userLists],
  );
  useEffect(() => listRef.current?.focus(), []);
  const selectedUserIndex = managedUserListIndex(selection, userLists);
  const selectedUser = userLists[selectedUserIndex];
  const selectedBuiltIn = managedBuiltInList(selection, builtInLists);
  const editingNewList = selection === 'new';

  const validateDraft = (
    rows: readonly ManagedSpreadsheetSortCustomList[],
    excludedId: number | null,
  ): readonly string[] | null => {
    const validation = parseSpreadsheetSortCustomList(text);
    if (!validation.ok) {
      setError(validation.message);
      return null;
    }
    const duplicate = [
      ...builtInLists.map((list) => list.entries),
      ...rows.filter((row) => row.id !== excludedId).map((row) => row.entries),
    ].some((entries) =>
      spreadsheetSortCustomListsEqual(entries, validation.entries),
    );
    if (duplicate) {
      setError('该自定义序列已存在。');
      return null;
    }
    setError(null);
    return Object.freeze([...validation.entries]);
  };

  const commitSelectedUser = (
    rows: readonly ManagedSpreadsheetSortCustomList[],
  ): readonly ManagedSpreadsheetSortCustomList[] | null => {
    const index = managedUserListIndex(selection, rows);
    const selected = rows[index];
    if (!selected) return rows;
    const entries = validateDraft(rows, selected.id);
    if (!entries) return null;
    return rows.map((row) =>
      row.id === selected.id ? { ...row, entries } : row,
    );
  };

  const select = (nextSelection: string) => {
    let rows = userLists;
    if (selectedUser) {
      const committed = commitSelectedUser(rows);
      if (!committed) return;
      rows = committed;
      setUserLists(rows);
    } else if (editingNewList && text.trim()) {
      setError('请先添加当前序列，或清空项目后再切换。');
      return;
    }
    setSelection(nextSelection);
    setText(managedSelectionText(nextSelection, builtInLists, rows));
    setError(null);
  };

  const saveSelectedUser = () => {
    const committed = commitSelectedUser(userLists);
    if (!committed) return;
    setUserLists(committed);
    setText(managedSelectionText(selection, builtInLists, committed));
  };

  const addNewList = () => {
    if (userLists.length >= MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS) {
      setError(
        `当前编辑器最多保留 ${MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS} 个自定义序列。`,
      );
      return;
    }
    const entries = validateDraft(userLists, null);
    if (!entries) return;
    const id = nextIdRef.current;
    nextIdRef.current += 1;
    const next = [
      ...userLists,
      { id, entries, originalEntries: null },
    ] as const;
    setUserLists(next);
    setSelection(`user:${id}`);
    setText(entries.join('\n'));
  };

  const moveSelectedUser = (offset: -1 | 1) => {
    const committed = commitSelectedUser(userLists);
    if (!committed) return;
    const index = managedUserListIndex(selection, committed);
    const nextIndex = index + offset;
    if (index < 0 || nextIndex < 0 || nextIndex >= committed.length) return;
    const next = committed.map(cloneManagedCustomList);
    const selected = next[index];
    const target = next[nextIndex];
    if (!selected || !target) return;
    next[index] = target;
    next[nextIndex] = selected;
    setUserLists(next);
    setText(selected.entries.join('\n'));
  };

  const deleteSelectedUser = () => {
    if (!selectedUser) return;
    const next = userLists.filter((row) => row.id !== selectedUser.id);
    setUserLists(next);
    const fallback = next[selectedUserIndex] ?? next[selectedUserIndex - 1];
    const nextSelection = fallback ? `user:${fallback.id}` : 'built-in:0';
    setSelection(nextSelection);
    setText(managedSelectionText(nextSelection, builtInLists, next));
    setError(null);
  };

  const apply = () => {
    let rows = userLists;
    if (selectedUser) {
      const committed = commitSelectedUser(rows);
      if (!committed) return;
      rows = committed;
    } else if (editingNewList && text.trim()) {
      if (rows.length >= MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS) {
        setError(
          `当前编辑器最多保留 ${MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS} 个自定义序列。`,
        );
        return;
      }
      const entries = validateDraft(rows, null);
      if (!entries) return;
      const id = nextIdRef.current;
      nextIdRef.current += 1;
      rows = [...rows, { id, entries, originalEntries: null }];
    }
    onApply(managedCustomListResult(initialUserLists, rows));
    onClose();
  };

  return (
    <Dialog
      title="自定义序列"
      description="管理排序时可复用的本地序列。内置月份和星期序列保持只读。"
      className="work-spreadsheet-sort-custom-list-manager"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button tone="primary" onClick={apply}>
            确定
          </Button>
        </>
      }
    >
      <div className="work-spreadsheet-sort-custom-list-manager-layout">
        <section aria-label="可用序列">
          <label>
            <span>自定义序列</span>
            <select
              ref={listRef}
              aria-label="自定义序列列表"
              size={10}
              value={selection}
              onChange={(event) => select(event.currentTarget.value)}
            >
              <optgroup label="内置序列">
                {builtInLists.map((list, index) => (
                  <option key={`built-in:${index}`} value={`built-in:${index}`}>
                    {list.label}
                  </option>
                ))}
              </optgroup>
              {userLists.length ? (
                <optgroup label="用户序列">
                  {userLists.map((list) => (
                    <option key={list.id} value={`user:${list.id}`}>
                      {userListLabels.get(list.id)}
                    </option>
                  ))}
                </optgroup>
              ) : null}
            </select>
          </label>
          <Button
            tone="quiet"
            type="button"
            disabled={
              userLists.length >= MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS
            }
            onClick={() => select('new')}
          >
            <ListPlus size={15} aria-hidden="true" />
            新建序列
          </Button>
        </section>

        <section aria-label="序列项目">
          <label>
            <span>序列项目（每行一个项目）</span>
            <textarea
              aria-label="自定义序列项目"
              aria-invalid={error ? true : undefined}
              readOnly={Boolean(selectedBuiltIn)}
              rows={12}
              value={text}
              onChange={(event) => {
                setText(event.currentTarget.value);
                setError(null);
              }}
            />
          </label>
          <p>
            {selectedBuiltIn
              ? '内置序列不可修改或删除。'
              : editingNewList
                ? '每行、英文逗号或中文逗号可分隔一个项目。'
                : '修改项目后保存；排序和删除在点击“确定”后统一提交。'}
          </p>
          {error ? (
            <p className="work-spreadsheet-sort-custom-list-error" role="alert">
              {error}
            </p>
          ) : null}
          <div className="work-spreadsheet-sort-custom-list-manager-actions">
            {editingNewList ? (
              <Button tone="primary" type="button" onClick={addNewList}>
                <ListPlus size={15} aria-hidden="true" />
                添加序列
              </Button>
            ) : selectedUser ? (
              <Button tone="primary" type="button" onClick={saveSelectedUser}>
                <Save size={15} aria-hidden="true" />
                保存更改
              </Button>
            ) : null}
            <Button
              tone="quiet"
              type="button"
              aria-label="上移序列"
              title="上移序列"
              disabled={selectedUserIndex <= 0}
              onClick={() => moveSelectedUser(-1)}
            >
              <ArrowUp size={15} aria-hidden="true" />
            </Button>
            <Button
              tone="quiet"
              type="button"
              aria-label="下移序列"
              title="下移序列"
              disabled={
                selectedUserIndex < 0 ||
                selectedUserIndex >= userLists.length - 1
              }
              onClick={() => moveSelectedUser(1)}
            >
              <ArrowDown size={15} aria-hidden="true" />
            </Button>
            <Button
              tone="quiet"
              type="button"
              aria-label="删除序列"
              title="删除序列"
              disabled={!selectedUser}
              onClick={deleteSelectedUser}
            >
              <Trash2 size={15} aria-hidden="true" />
            </Button>
          </div>
        </section>
      </div>
      <p className="work-spreadsheet-sort-custom-list-manager-note">
        最多保存 {MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS}{' '}
        个用户序列。修改或删除当前使用的序列会同步更新对应排序条件。
      </p>
    </Dialog>
  );
}

function managedCustomListResult(
  initial: readonly ManagedSpreadsheetSortCustomList[],
  current: readonly ManagedSpreadsheetSortCustomList[],
): SpreadsheetSortCustomListManagementResult {
  const changes: SpreadsheetSortCustomListManagementChange[] = [];
  for (const original of initial) {
    if (!original.originalEntries) continue;
    const next = current.find((row) => row.id === original.id);
    if (!next) {
      changes.push({ previous: original.originalEntries, next: null });
    } else if (
      !spreadsheetSortCustomListsEqual(original.originalEntries, next.entries)
    ) {
      changes.push({
        previous: original.originalEntries,
        next: Object.freeze([...next.entries]),
      });
    }
  }
  return Object.freeze({
    lists: Object.freeze(current.map((row) => Object.freeze([...row.entries]))),
    changes: Object.freeze(changes),
  });
}

function managedUserListIndex(
  selection: string,
  lists: readonly ManagedSpreadsheetSortCustomList[],
): number {
  if (!selection.startsWith('user:')) return -1;
  const id = Number(selection.slice('user:'.length));
  return Number.isSafeInteger(id)
    ? lists.findIndex((list) => list.id === id)
    : -1;
}

function managedBuiltInList(
  selection: string,
  lists: readonly SpreadsheetSortCustomList[],
): SpreadsheetSortCustomList | null {
  if (!selection.startsWith('built-in:')) return null;
  const index = Number(selection.slice('built-in:'.length));
  return Number.isSafeInteger(index) ? (lists[index] ?? null) : null;
}

function managedSelectionText(
  selection: string,
  builtInLists: readonly SpreadsheetSortCustomList[],
  userLists: readonly ManagedSpreadsheetSortCustomList[],
): string {
  return (
    managedBuiltInList(selection, builtInLists)?.entries ??
    userLists[managedUserListIndex(selection, userLists)]?.entries ??
    []
  ).join('\n');
}

function cloneManagedCustomList(
  list: ManagedSpreadsheetSortCustomList,
): ManagedSpreadsheetSortCustomList {
  return {
    ...list,
    entries: Object.freeze([...list.entries]),
    originalEntries: list.originalEntries
      ? Object.freeze([...list.originalEntries])
      : null,
  };
}

function managedCustomListLabel(entries: readonly string[]): string {
  return (
    createSpreadsheetSortCustomList(entries, 'session')?.label ??
    entries.join(' → ')
  );
}
