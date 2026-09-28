import {
  ArrowDown,
  ArrowUp,
  ListOrdered,
  ListPlus,
  Settings2,
  Trash2,
} from 'lucide-react';
import { type FormEvent, useId, useMemo, useRef, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeCheckbox, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  MAX_SPREADSHEET_SORT_KEYS,
  type SpreadsheetSortDialogSource,
  type SpreadsheetSortDialogValue,
  type SpreadsheetSortKey,
  type SpreadsheetSortOptions,
} from './spreadsheet-sort';
import {
  type SpreadsheetSortAppearanceField,
  type SpreadsheetSortAppearanceTarget,
  spreadsheetSortAppearanceFields,
  spreadsheetSortAppearanceTargets,
  spreadsheetSortAppearanceTargetsEqual,
} from './spreadsheet-sort-appearance';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS,
  mergeSpreadsheetSortCustomLists,
  parseSpreadsheetSortCustomList,
  type SpreadsheetSortCustomList,
  spreadsheetSortCustomListsEqual,
} from './spreadsheet-sort-custom-list';
import { SpreadsheetSortCustomListEditor } from './spreadsheet-sort-custom-list-editor';
import {
  type SpreadsheetSortCustomListManagementResult,
  SpreadsheetSortCustomListManagerDialog,
} from './spreadsheet-sort-custom-list-manager';
import { SpreadsheetSortOptionsDialog } from './spreadsheet-sort-options-dialog';
import {
  nextSpreadsheetSortKey,
  SpreadsheetSortOrderControls,
  spreadsheetSortAppearanceKey,
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
  onUpdateCustomLists,
  onClose,
}: {
  source: SpreadsheetSortDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetSortDialogValue) => boolean;
  onRememberCustomList?: (
    list: SpreadsheetSortCustomList,
  ) => SpreadsheetSortCustomList | undefined;
  onUpdateCustomLists?: (
    lists: readonly (readonly string[])[],
  ) => readonly SpreadsheetSortCustomList[] | undefined;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const [value, setValue] = useState<SpreadsheetSortDialogValue>(() => ({
    hasHeader: source.value.hasHeader,
    keys: source.value.keys.map(cloneSpreadsheetSortKey),
    orientation: source.value.orientation,
    caseSensitive: source.value.caseSensitive,
    textMethod: source.value.textMethod,
  }));
  const [verticalHasHeader, setVerticalHasHeader] = useState(
    source.value.hasHeader,
  );
  const structuralScope = source.scope !== undefined;
  const [optionsOpen, setOptionsOpen] = useState(false);
  const [customListManagerOpen, setCustomListManagerOpen] = useState(false);
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
  const customListManagerButtonRef = useRef<HTMLButtonElement>(null);
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
        caseSensitive: value.caseSensitive,
        textMethod: value.textMethod,
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
      const userListCount = customLists.filter(
        (list) => list.source !== 'built-in',
      ).length;
      if (userListCount >= MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS) {
        setCustomListDraft((current) =>
          current
            ? {
                ...current,
                error: officeMessage(messages, 'spreadsheet.sort.customListLimit', {
                  n: String(MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS),
                }),
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
      const remembered = onRememberCustomList?.(created);
      selected = normalizedRememberedCustomList(remembered, created);
      setCustomLists((current) => Object.freeze([...current, selected!]));
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

  const applyOptions = (options: SpreadsheetSortOptions) => {
    setOptionsOpen(false);
    if (structuralScope && options.orientation !== 'top-to-bottom') return;
    if (options.orientation === value.orientation) {
      setValue((current) => ({
        ...current,
        caseSensitive: options.caseSensitive,
        textMethod: options.textMethod,
      }));
      return;
    }
    const nextFields =
      options.orientation === 'top-to-bottom' ? source.columns : source.rows;
    const preferredIndex =
      options.orientation === 'top-to-bottom'
        ? (source.value.keys[0]?.index ?? nextFields[0]?.index)
        : source.activeRow;
    const index =
      nextFields.find((field) => field.index === preferredIndex)?.index ??
      nextFields[0]?.index;
    if (index === undefined) return;
    setCustomListDraft(null);
    setValue({
      ...options,
      hasHeader:
        options.orientation === 'top-to-bottom' ? verticalHasHeader : false,
      keys: [{ index, direction: 'ascending' }],
    });
  };

  const applyCustomListManagement = (
    result: SpreadsheetSortCustomListManagementResult,
  ) => {
    const currentUserLists = customLists.filter(
      (list) => list.source !== 'built-in',
    );
    const preferencesChanged =
      currentUserLists.length !== result.lists.length ||
      result.lists.some(
        (entries, index) =>
          !spreadsheetSortCustomListsEqual(
            currentUserLists[index]?.entries ?? [],
            entries,
          ),
      );
    const fallback = result.lists
      .map((entries) => createSpreadsheetSortCustomList(entries, 'session'))
      .filter((list): list is SpreadsheetSortCustomList => list !== null);
    const updated = preferencesChanged
      ? (onUpdateCustomLists?.(result.lists) ?? fallback)
      : currentUserLists;
    setCustomLists(mergeSpreadsheetSortCustomLists(updated));
    setCustomListDraft(null);
    setValue((current) => ({
      ...current,
      keys: current.keys.map((key) =>
        reconcileManagedCustomListKey(key, result.changes),
      ),
    }));
  };

  return (
    <>
      <Dialog
        title={officeMessage(messages, 'spreadsheet.sort.title')}
        description={`${source.sheetName}!${source.rangeReference}`}
        className="work-spreadsheet-sort-dialog"
        restoreFocusTarget={restoreFocusTarget}
        onClose={onClose}
        footer={
          <>
            <Button tone="quiet" onClick={onClose}>
              {officeMessage(messages, 'spreadsheet.sort.cancel')}
            </Button>
            <Button tone="primary" type="submit" form={formId}>
              {officeMessage(messages, 'spreadsheet.sort.ok')}
            </Button>
          </>
        }
      >
        <form id={formId} onSubmit={submit}>
          <div className="work-spreadsheet-sort-toolbar">
            <div className="work-spreadsheet-sort-toolbar-actions">
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
                {officeMessage(messages, 'spreadsheet.sort.addLevel')}
              </Button>
              <Button
                ref={optionsButtonRef}
                tone="quiet"
                type="button"
                onClick={() => setOptionsOpen(true)}
              >
                <Settings2 size={15} aria-hidden="true" />
                {officeMessage(messages, 'spreadsheet.sort.options')}
              </Button>
              <Button
                ref={customListManagerButtonRef}
                tone="quiet"
                type="button"
                aria-label={officeMessage(messages, 'spreadsheet.sort.manageListsAria')}
                title={officeMessage(messages, 'spreadsheet.sort.manageListsAria')}
                onClick={() => setCustomListManagerOpen(true)}
              >
                <ListOrdered size={15} aria-hidden="true" />
                {officeMessage(messages, 'spreadsheet.sort.customLists')}
              </Button>
            </div>
            <OfficeCheckbox
              ariaLabel={officeMessage(messages, 'spreadsheet.sort.hasHeader')}
              checked={value.hasHeader}
              disabled={
                structuralScope || value.orientation === 'left-to-right'
              }
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
              {officeMessage(messages, 'spreadsheet.sort.hasHeader')}
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
                    {index === 0
                      ? officeMessage(messages, 'spreadsheet.sort.primaryKey')
                      : officeMessage(messages, 'spreadsheet.sort.secondaryKey', {
                          n: String(index),
                        })}
                  </legend>
                  <div className="work-office-field">
                    <span>
                      {value.orientation === 'top-to-bottom'
                        ? officeMessage(messages, 'spreadsheet.sort.column')
                        : officeMessage(messages, 'spreadsheet.sort.row')}
                    </span>
                    <OfficeSelect
                      ariaLabel={officeMessage(
                        messages,
                        'spreadsheet.sort.levelAxisAria',
                        {
                          level: String(level),
                          axis:
                            value.orientation === 'top-to-bottom'
                              ? officeMessage(messages, 'spreadsheet.sort.column')
                              : officeMessage(messages, 'spreadsheet.sort.row'),
                        },
                      )}
                      value={String(key.index)}
                      options={fields.map((field) => ({
                        value: String(field.index),
                        label: field.label,
                      }))}
                      onValueChange={(raw) => {
                        const fieldIndex = Number(raw);
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
                    />
                  </div>
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
                      aria-label={officeMessage(messages, 'spreadsheet.sort.moveUpAria', { level: String(level) })}
                      title={officeMessage(messages, 'spreadsheet.sort.moveUpTitle')}
                      disabled={index === 0}
                      onClick={() => moveKey(index, -1)}
                    >
                      <ArrowUp size={15} aria-hidden="true" />
                    </Button>
                    <Button
                      tone="quiet"
                      type="button"
                      aria-label={officeMessage(messages, 'spreadsheet.sort.moveDownAria', { level: String(level) })}
                      title={officeMessage(messages, 'spreadsheet.sort.moveDownTitle')}
                      disabled={index === value.keys.length - 1}
                      onClick={() => moveKey(index, 1)}
                    >
                      <ArrowDown size={15} aria-hidden="true" />
                    </Button>
                    <Button
                      tone="quiet"
                      type="button"
                      aria-label={officeMessage(messages, 'spreadsheet.sort.deleteLevelAria', { level: String(level) })}
                      title={officeMessage(messages, 'spreadsheet.sort.deleteLevelTitle')}
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
                        {officeMessage(messages, 'spreadsheet.sort.editList')}
                      </Button>
                    </div>
                  ) : null}
                </fieldset>
              );
            })}
          </div>
          <p className="work-spreadsheet-sort-note">
            {source.scope?.kind === 'auto-filter'
              ? officeMessage(messages, 'spreadsheet.sort.hint.autoFilter')
              : source.scope?.kind === 'table'
                ? officeMessage(messages, 'spreadsheet.sort.hint.table')
                : officeMessage(messages, 'spreadsheet.sort.hint.default')}
          </p>
        </form>
      </Dialog>
      {optionsOpen ? (
        <SpreadsheetSortOptionsDialog
          value={{
            orientation: value.orientation,
            caseSensitive: value.caseSensitive,
            textMethod: value.textMethod,
          }}
          orientationLocked={structuralScope}
          restoreFocusTarget={() => optionsButtonRef.current}
          onApply={applyOptions}
          onClose={() => setOptionsOpen(false)}
        />
      ) : null}
      {customListManagerOpen ? (
        <SpreadsheetSortCustomListManagerDialog
          customLists={customLists}
          restoreFocusTarget={() => customListManagerButtonRef.current}
          onApply={applyCustomListManagement}
          onClose={() => setCustomListManagerOpen(false)}
        />
      ) : null}
    </>
  );
}

function reconcileManagedCustomListKey(
  key: SpreadsheetSortKey,
  changes: SpreadsheetSortCustomListManagementResult['changes'],
): SpreadsheetSortKey {
  if (key.customList === undefined) return key;
  const change = changes.find((candidate) =>
    spreadsheetSortCustomListsEqual(candidate.previous, key.customList ?? []),
  );
  if (!change) return key;
  return change.next
    ? { index: key.index, customList: [...change.next] }
    : { index: key.index, direction: 'ascending' };
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
  const candidates: SpreadsheetSortCustomList[] = [...source.customLists];
  for (const key of source.value.keys) {
    if (key.customList === undefined) continue;
    const list = createSpreadsheetSortCustomList(key.customList, 'session');
    if (list) candidates.push(list);
  }
  return mergeSpreadsheetSortCustomLists(candidates);
}

function normalizedRememberedCustomList(
  candidate: unknown,
  fallback: SpreadsheetSortCustomList,
): SpreadsheetSortCustomList {
  if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate)) {
    return fallback;
  }
  const source = (candidate as Partial<SpreadsheetSortCustomList>).source;
  const entries = (candidate as Partial<SpreadsheetSortCustomList>).entries;
  if (
    (source !== 'stored' && source !== 'session') ||
    !Array.isArray(entries)
  ) {
    return fallback;
  }
  return createSpreadsheetSortCustomList(entries, source) ?? fallback;
}
