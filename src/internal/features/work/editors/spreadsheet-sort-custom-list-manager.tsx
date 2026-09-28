import { ArrowDown, ArrowUp, ListPlus, Save, Trash2 } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, Dialog } from '../../../design-system/primitives';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS,
  parseSpreadsheetSortCustomList,
  spreadsheetSortCustomListsEqual,
  type SpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';
import { useOfficeMessages } from './office-messages-context';

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
  const messages = useOfficeMessages();
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
      setError(officeMessage(messages, 'spreadsheet.sort.listManager.duplicate'));
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
      setError(officeMessage(messages, 'spreadsheet.sort.listManager.commitOrClear'));
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
        officeMessage(messages, 'spreadsheet.sort.customListLimit', {
          n: String(MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS),
        }),
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
          officeMessage(messages, 'spreadsheet.sort.customListLimit', {
            n: String(MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS),
          }),
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
      title={officeMessage(messages, 'spreadsheet.sort.listManager.title')}
      description={officeMessage(messages, 'spreadsheet.sort.listManager.desc')}
      className="work-spreadsheet-sort-custom-list-manager"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.sort.cancel')}
          </Button>
          <Button tone="primary" onClick={apply}>
            {officeMessage(messages, 'spreadsheet.sort.ok')}
          </Button>
        </>
      }
    >
      <div className="work-spreadsheet-sort-custom-list-manager-layout">
        <section aria-label={officeMessage(messages, 'spreadsheet.sort.listManager.availableAria')}>
          <label>
            <span>{officeMessage(messages, 'spreadsheet.sort.listManager.listLabel')}</span>
            <select
              ref={listRef}
              aria-label={officeMessage(messages, 'spreadsheet.sort.listManager.listAria')}
              size={10}
              value={selection}
              onChange={(event) => select(event.currentTarget.value)}
            >
              <optgroup label={officeMessage(messages, 'spreadsheet.sort.listManager.groupBuiltin')}>
                {builtInLists.map((list, index) => (
                  <option key={`built-in:${index}`} value={`built-in:${index}`}>
                    {list.label}
                  </option>
                ))}
              </optgroup>
              {userLists.length ? (
                <optgroup label={officeMessage(messages, 'spreadsheet.sort.listManager.groupUser')}>
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
            {officeMessage(messages, 'spreadsheet.sort.listManager.newList')}
          </Button>
        </section>

        <section aria-label={officeMessage(messages, 'spreadsheet.sort.listManager.entriesAria')}>
          <label>
            <span>{officeMessage(messages, 'spreadsheet.sort.listManager.entriesLabel')}</span>
            <textarea
              aria-label={officeMessage(messages, 'spreadsheet.sort.listManager.entriesFieldAria')}
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
              ? officeMessage(messages, 'spreadsheet.sort.listManager.builtinReadonly')
              : editingNewList
                ? officeMessage(messages, 'spreadsheet.sort.listManager.entriesHint')
                : officeMessage(messages, 'spreadsheet.sort.listManager.saveHint')}
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
                {officeMessage(messages, 'spreadsheet.sort.listManager.add')}
              </Button>
            ) : selectedUser ? (
              <Button tone="primary" type="button" onClick={saveSelectedUser}>
                <Save size={15} aria-hidden="true" />
                {officeMessage(messages, 'spreadsheet.sort.listManager.save')}
              </Button>
            ) : null}
            <Button
              tone="quiet"
              type="button"
              aria-label={officeMessage(messages, 'spreadsheet.sort.listManager.moveUpAria')}
              title={officeMessage(messages, 'spreadsheet.sort.listManager.moveUpTitle')}
              disabled={selectedUserIndex <= 0}
              onClick={() => moveSelectedUser(-1)}
            >
              <ArrowUp size={15} aria-hidden="true" />
            </Button>
            <Button
              tone="quiet"
              type="button"
              aria-label={officeMessage(messages, 'spreadsheet.sort.listManager.moveDownAria')}
              title={officeMessage(messages, 'spreadsheet.sort.listManager.moveDownTitle')}
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
              aria-label={officeMessage(messages, 'spreadsheet.sort.listManager.deleteAria')}
              title={officeMessage(messages, 'spreadsheet.sort.listManager.deleteTitle')}
              disabled={!selectedUser}
              onClick={deleteSelectedUser}
            >
              <Trash2 size={15} aria-hidden="true" />
            </Button>
          </div>
        </section>
      </div>
      <p className="work-spreadsheet-sort-custom-list-manager-note">
        {officeMessage(messages, 'spreadsheet.sort.listManager.footerLimit', {
          n: String(MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS),
        })}
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
