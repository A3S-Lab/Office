import type { Selection } from '@fortune-sheet/core';
import { Plus, RefreshCw, Trash2 } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import {
  Button,
  CollectionState,
  InlineNotice,
  StateView,
} from '../../../design-system/primitives';
import {
  WORK_SPREADSHEET_DEFAULT_PIVOT_STYLE,
  WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS,
} from '../work-spreadsheet-pivot-styles';
import {
  spreadsheetPivotFilterValueKey,
  spreadsheetPivotReportFilterSelection,
} from '../work-spreadsheet-pivot-values';
import {
  createSpreadsheetPivotFromSelection,
  defaultPivotValueCaption,
  deleteSpreadsheetPivotTable,
  refreshSpreadsheetPivotTables,
  type SpreadsheetPivotFilterItem,
  spreadsheetPivotAggregationLabel,
  spreadsheetPivotFields,
  spreadsheetPivotFilterItems,
  spreadsheetPivotValidation,
} from '../work-spreadsheet-pivots';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetPivotAggregation,
  WorkSpreadsheetPivotFilterValue,
  WorkSpreadsheetPivotTable,
} from '../work-types';
import {
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { useOfficeDraft } from './use-office-draft';

interface SpreadsheetPivotPanelProps {
  content: WorkSpreadsheetContent;
  activeSheetId: string;
  selection?: Selection;
  onChange: (content: WorkSpreadsheetContent) => void;
}

interface PivotListItem {
  ownerSheetId: string;
  ownerSheetName: string;
  pivot: WorkSpreadsheetPivotTable;
}

type PivotDraft = WorkSpreadsheetPivotTable & { ownerSheetId: string };
type PivotFieldRole = 'unused' | 'row' | 'column' | 'filter' | 'value';

const AGGREGATIONS: WorkSpreadsheetPivotAggregation[] = [
  'sum',
  'count',
  'counta',
  'average',
  'max',
  'min',
  'product',
  'stdDev',
  'stdDevP',
  'var',
  'varP',
];

export function SpreadsheetPivotPanel({
  content,
  activeSheetId,
  selection,
  onChange,
}: SpreadsheetPivotPanelProps) {
  const messages = useOfficeMessages();
  const items = useMemo(
    () =>
      content.sheets.flatMap((sheet) =>
        (sheet.pivotTables ?? []).flatMap((pivot) =>
          sheet.id
            ? [{ ownerSheetId: sheet.id, ownerSheetName: sheet.name, pivot }]
            : [],
        ),
      ),
    [content.sheets],
  );
  const [selectedKey, setSelectedKey] = useState<string | null>(() =>
    pivotKey(items[0]),
  );
  const {
    cancelDraft: resetDraft,
    dirty,
    draft,
    draftRef,
    replaceDraft,
    setDraft,
    syncDraft,
  } = useOfficeDraft<PivotDraft | null>(() =>
    items[0] ? pivotDraft(items[0]) : null,
  );
  const [error, setError] = useState('');
  const optimisticPivotKey = useRef<string | null>(null);

  useEffect(() => {
    if (!items.length) {
      if (selectedKey && optimisticPivotKey.current === selectedKey) return;
      setSelectedKey(null);
      replaceDraft(null);
      return;
    }
    const selected = items.find((item) => pivotKey(item) === selectedKey);
    if (selected) {
      if (optimisticPivotKey.current === selectedKey)
        optimisticPivotKey.current = null;
      syncDraft(
        pivotDraft(selected),
        pivotDraftKey(draftRef.current) !== selectedKey,
      );
      return;
    }
    const first = items[0];
    setSelectedKey(pivotKey(first));
    replaceDraft(pivotDraft(first));
  }, [draftRef, items, replaceDraft, syncDraft]);

  const fields = useMemo(
    () => (draft ? spreadsheetPivotFields(content, draft) : []),
    [content, draft],
  );
  const selectPivot = (item: PivotListItem) => {
    const nextKey = pivotKey(item);
    if (nextKey === selectedKey) return;
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.pivot.error.unsaved'));
      return;
    }
    setSelectedKey(pivotKey(item));
    replaceDraft(pivotDraft(item));
    setError('');
  };
  const addPivot = () => {
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.pivot.error.unsaved'));
      return;
    }
    if (!selection) {
      setError(officeMessage(messages, 'spreadsheet.pivot.error.needSelection'));
      return;
    }
    const created = createSpreadsheetPivotFromSelection(
      content,
      activeSheetId,
      selection,
    );
    if (created.error || !created.ownerSheetId || !created.pivotId) {
      setError(created.error ?? officeMessage(messages, 'spreadsheet.pivot.error.createFailed'));
      return;
    }
    onChange(created.content);
    const createdKey = `${created.ownerSheetId}:${created.pivotId}`;
    optimisticPivotKey.current = createdKey;
    setSelectedKey(createdKey);
    const owner = created.content.sheets.find(
      (sheet) => sheet.id === created.ownerSheetId,
    );
    const pivot = owner?.pivotTables?.find(
      (candidate) => candidate.id === created.pivotId,
    );
    replaceDraft(
      owner && pivot
        ? pivotDraft({
            ownerSheetId: owner.id!,
            ownerSheetName: owner.name,
            pivot,
          })
        : null,
    );
    setError('');
  };
  const savePivot = () => {
    if (!draft) return;
    const name = draft.name.trim();
    if (!/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) || name.length > 255) {
      setError(officeMessage(messages, 'spreadsheet.pivot.error.nameInvalid'));
      return;
    }
    if (
      items.some(
        (item) =>
          item.pivot.id !== draft.id &&
          item.pivot.name.trim().toLocaleLowerCase() ===
            name.toLocaleLowerCase(),
      )
    ) {
      setError(officeMessage(messages, 'spreadsheet.pivot.error.nameDuplicate'));
      return;
    }
    const saved: WorkSpreadsheetPivotTable = {
      ...draft,
      name,
      sourceReference: draft.sourceReference.trim().replace(/^=/, ''),
      anchor: draft.anchor.trim().replace(/^=/, ''),
      styleName: draft.styleName || WORK_SPREADSHEET_DEFAULT_PIVOT_STYLE,
      rowFields: [...draft.rowFields].sort((left, right) => left - right),
      columnFields: [...draft.columnFields].sort((left, right) => left - right),
      reportFilters: (draft.reportFilters ?? [])
        .map((filter) => ({ ...filter }))
        .sort((left, right) => left.fieldIndex - right.fieldIndex),
      values: draft.values
        .map((value) => ({
          ...value,
          caption:
            value.caption?.trim() ||
            defaultPivotValueCaption(
              fields[value.fieldIndex]?.name ??
                officeMessage(messages, 'spreadsheet.pivot.field.fallback', {
                  n: String(value.fieldIndex + 1),
                }),
              value.aggregation,
            ),
        }))
        .sort((left, right) => left.fieldIndex - right.fieldIndex),
    };
    const sheets = content.sheets.map((sheet) =>
      sheet.id === draft.ownerSheetId
        ? {
            ...sheet,
            pivotTables: (sheet.pivotTables ?? []).map((pivot) =>
              pivot.id === saved.id ? saved : pivot,
            ),
          }
        : sheet,
    );
    const candidate = { ...content, sheets };
    const validation = spreadsheetPivotValidation(
      candidate,
      draft.ownerSheetId,
      saved,
    );
    if (!validation.valid) {
      setError(validation.message ?? officeMessage(messages, 'spreadsheet.pivot.error.invalid'));
      return;
    }
    const refreshed = refreshSpreadsheetPivotTables(candidate);
    onChange(refreshed);
    const refreshedOwner = refreshed.sheets.find(
      (sheet) => sheet.id === draft.ownerSheetId,
    );
    const refreshedPivot = refreshedOwner?.pivotTables?.find(
      (pivot) => pivot.id === saved.id,
    );
    if (refreshedOwner && refreshedPivot) {
      replaceDraft(
        pivotDraft({
          ownerSheetId: refreshedOwner.id!,
          ownerSheetName: refreshedOwner.name,
          pivot: refreshedPivot,
        }),
      );
    }
    setError('');
  };
  const cancelDraft = () => {
    resetDraft();
    setError('');
  };
  const deletePivot = () => {
    if (!draft) return;
    optimisticPivotKey.current = null;
    const next = deleteSpreadsheetPivotTable(
      content,
      draft.ownerSheetId,
      draft.id,
    );
    onChange(next);
    const remaining = items.find(
      (item) => pivotKey(item) !== `${draft.ownerSheetId}:${draft.id}`,
    );
    setSelectedKey(pivotKey(remaining));
    replaceDraft(remaining ? pivotDraft(remaining) : null);
    setError('');
  };
  const refreshAll = () => {
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.pivot.error.unsaved'));
      return;
    }
    onChange(refreshSpreadsheetPivotTables(content));
    setError('');
  };

  return (
    <fieldset
      className="work-spreadsheet-pivot-manager"
      data-office-escape-consumer={dirty || undefined}
      onKeyDown={(event) => {
        if (event.key !== 'Escape' || event.defaultPrevented || !dirty) return;
        event.preventDefault();
        event.stopPropagation();
        cancelDraft();
      }}
    >
      <legend className="sr-only">{officeMessage(messages, 'spreadsheet.pivot.legend')}</legend>
      <aside aria-label={officeMessage(messages, 'spreadsheet.pivot.listAria')}>
        <Button className="create" tone="secondary" onClick={addPivot}>
          <Plus size={13} />
          {officeMessage(messages, 'spreadsheet.pivot.newFromSelection')}
        </Button>
        <Button
          className="refresh"
          tone="secondary"
          disabled={!items.length}
          onClick={refreshAll}
        >
          <RefreshCw size={13} />
          {officeMessage(messages, 'spreadsheet.pivot.refreshAll')}
        </Button>
        <div className="work-spreadsheet-pivot-list">
          {items.map((item) => (
            <button
              type="button"
              className={pivotKey(item) === selectedKey ? 'active' : ''}
              key={pivotKey(item)}
              onClick={() => selectPivot(item)}
            >
              <strong>{item.pivot.name}</strong>
              <span>
                {item.ownerSheetName} ·{' '}
                {item.pivot.outputReference ?? item.pivot.anchor}
              </span>
            </button>
          ))}
          {!items.length && (
            <CollectionState
              className="work-office-collection-empty"
              role="status"
            >
              {officeMessage(messages, 'spreadsheet.pivot.empty')}
            </CollectionState>
          )}
        </div>
      </aside>
      {draft ? (
        <form
          onSubmit={(event) => {
            event.preventDefault();
            savePivot();
          }}
        >
          <div className="work-spreadsheet-pivot-fields">
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.pivot.name')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.pivot.nameAria')}
                value={draft.name}
                maxLength={255}
                onChange={(event) =>
                  setDraft({ ...draft, name: event.target.value })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.pivot.sourceSheet')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'spreadsheet.pivot.sourceSheetAria')}
                value={draft.sourceSheetId}
                options={content.sheets.flatMap((sheet) =>
                  sheet.id ? [{ value: sheet.id, label: sheet.name }] : [],
                )}
                onValueChange={(sourceSheetId) =>
                  setDraft({ ...draft, sourceSheetId })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.pivot.sourceRange')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.pivot.sourceRangeAria')}
                value={draft.sourceReference}
                placeholder="A1:D200"
                onChange={(event) =>
                  setDraft({ ...draft, sourceReference: event.target.value })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.pivot.output')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.pivot.outputAria')}
                value={draft.anchor}
                placeholder="A1"
                onChange={(event) =>
                  setDraft({ ...draft, anchor: event.target.value })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.pivot.style')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'spreadsheet.pivot.styleAria')}
                value={draft.styleName}
                options={[...WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS]}
                onValueChange={(styleName) => setDraft({ ...draft, styleName })}
              />
            </div>
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(messages, 'spreadsheet.pivot.refreshOnOpen')}
              checked={draft.refreshOnLoad}
              onCheckedChange={(refreshOnLoad) =>
                setDraft({ ...draft, refreshOnLoad })
              }
            >
              {officeMessage(messages, 'spreadsheet.pivot.refreshOnOpen')}
            </OfficeCheckbox>
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(messages, 'spreadsheet.pivot.showColGrand')}
              checked={draft.rowGrandTotals}
              onCheckedChange={(rowGrandTotals) =>
                setDraft({ ...draft, rowGrandTotals })
              }
            >
              {officeMessage(messages, 'spreadsheet.pivot.showColGrand')}
            </OfficeCheckbox>
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(messages, 'spreadsheet.pivot.showRowGrand')}
              checked={draft.columnGrandTotals}
              onCheckedChange={(columnGrandTotals) =>
                setDraft({ ...draft, columnGrandTotals })
              }
            >
              {officeMessage(messages, 'spreadsheet.pivot.showRowGrand')}
            </OfficeCheckbox>
          </div>
          <section
            className="work-spreadsheet-pivot-layout"
            aria-label={officeMessage(messages, 'spreadsheet.pivot.layoutAria')}
          >
            <header>
              <strong>{officeMessage(messages, 'spreadsheet.pivot.layoutTitle')}</strong>
              <span>{officeMessage(messages, 'spreadsheet.pivot.layoutHint')}</span>
            </header>
            {fields.length ? (
              <div>
                {fields.map((field) => {
                  const role = pivotFieldRole(draft, field.index);
                  const value = draft.values.find(
                    (candidate) => candidate.fieldIndex === field.index,
                  );
                  const filter = draft.reportFilters?.find(
                    (candidate) => candidate.fieldIndex === field.index,
                  );
                  const filterItems =
                    role === 'filter'
                      ? spreadsheetPivotFilterItems(content, draft, field.index)
                      : [];
                  return (
                    <div
                      className="work-spreadsheet-pivot-field-row"
                      key={field.index}
                    >
                      <strong>{field.name}</strong>
                      <span>{field.numeric
                        ? officeMessage(messages, 'spreadsheet.pivot.field.numeric')
                        : officeMessage(messages, 'spreadsheet.pivot.field.category')}</span>
                      <OfficeSelect
                        ariaLabel={officeMessage(messages, 'spreadsheet.pivot.field.areaAria', { name: field.name })}
                        value={role}
                        options={[
                          { value: 'unused', label: officeMessage(messages, 'spreadsheet.pivot.area.unused') },
                          { value: 'row', label: officeMessage(messages, 'spreadsheet.pivot.area.row') },
                          { value: 'column', label: officeMessage(messages, 'spreadsheet.pivot.area.column') },
                          { value: 'filter', label: officeMessage(messages, 'spreadsheet.pivot.area.filter') },
                          { value: 'value', label: officeMessage(messages, 'spreadsheet.pivot.area.value') },
                        ]}
                        onValueChange={(nextRole) =>
                          setDraft(
                            assignPivotFieldRole(
                              draft,
                              field.index,
                              nextRole as PivotFieldRole,
                              field.name,
                              field.numeric,
                            ),
                          )
                        }
                      />
                      {role === 'value' && value ? (
                        <>
                          <OfficeSelect
                            ariaLabel={officeMessage(messages, 'spreadsheet.pivot.aggAria', { name: field.name })}
                            value={value.aggregation}
                            options={AGGREGATIONS.map((aggregation) => ({
                              value: aggregation,
                              label:
                                spreadsheetPivotAggregationLabel(aggregation),
                            }))}
                            onValueChange={(aggregation) =>
                              setDraft(
                                updatePivotValue(draft, field.index, {
                                  aggregation:
                                    aggregation as WorkSpreadsheetPivotAggregation,
                                }),
                              )
                            }
                          />
                          <OfficeTextField
                            aria-label={officeMessage(messages, 'spreadsheet.pivot.valueCaptionAria', { name: field.name })}
                            value={value.caption ?? ''}
                            placeholder={defaultPivotValueCaption(
                              field.name,
                              value.aggregation,
                            )}
                            onChange={(event) =>
                              setDraft(
                                updatePivotValue(draft, field.index, {
                                  caption: event.target.value,
                                }),
                              )
                            }
                          />
                        </>
                      ) : role === 'filter' && filter ? (
                        <>
                          <fieldset
                            className="pivot-slicer-filter"
                            aria-label={`${field.name} slicer filter`}
                          >
                            <OfficeCheckbox
                              className="check"
                              ariaLabel={officeMessage(messages, 'spreadsheet.pivot.filterAllAria', { name: field.name })}
                              checked={
                                spreadsheetPivotReportFilterSelection(filter)
                                  .kind === 'all'
                              }
                              onCheckedChange={(checked) =>
                                setDraft(
                                  updatePivotReportFilterItems(
                                    draft,
                                    field.index,
                                    checked ? undefined : [],
                                  ),
                                )
                              }
                            >
                              {officeMessage(messages, 'spreadsheet.pivot.filterAll')}
                            </OfficeCheckbox>
                            {filterItems.map((item) => {
                              const selection =
                                spreadsheetPivotReportFilterSelection(filter);
                              const itemKey = spreadsheetPivotFilterValueKey(
                                item.value,
                              );
                              const checked =
                                selection.kind === 'all' ||
                                (selection.kind === 'items' &&
                                  selection.items.some(
                                    (value) =>
                                      spreadsheetPivotFilterValueKey(value) ===
                                      itemKey,
                                  ));
                              return (
                                <OfficeCheckbox
                                  key={itemKey}
                                  className="check"
                                  ariaLabel={`${field.name} ${item.label}`}
                                  checked={checked}
                                  onCheckedChange={(nextChecked) =>
                                    setDraft(
                                      togglePivotReportFilterItem(
                                        draft,
                                        field.index,
                                        item.value,
                                        nextChecked,
                                        filterItems,
                                      ),
                                    )
                                  }
                                >
                                  {item.label}
                                </OfficeCheckbox>
                              );
                            })}
                          </fieldset>
                          <span className="filter-hint">{officeMessage(messages, 'spreadsheet.pivot.filterHint')}</span>
                        </>
                      ) : (
                        <span className="placeholder">—</span>
                      )}
                    </div>
                  );
                })}
              </div>
            ) : (
              <p>{officeMessage(messages, 'spreadsheet.pivot.configHint')}</p>
            )}
          </section>
          <div className="actions">
            <Button tone="danger" onClick={deletePivot}>
              <Trash2 size={13} />
              {officeMessage(messages, 'spreadsheet.pivot.delete')}
            </Button>
            {error && (
              <InlineNotice
                className="work-office-form-error"
                tone="danger"
                role="alert"
              >
                {error}
              </InlineNotice>
            )}
            <Button tone="secondary" disabled={!dirty} onClick={cancelDraft}>
              {officeMessage(messages, 'spreadsheet.pivot.cancel')}
            </Button>
            <Button type="submit" tone="primary" disabled={!dirty}>
              {officeMessage(messages, 'spreadsheet.pivot.save')}
            </Button>
          </div>
        </form>
      ) : (
        <StateView
          className="work-spreadsheet-pivot-empty"
          size="compact"
          title={officeMessage(messages, 'spreadsheet.pivot.createTitle')}
          description={officeMessage(messages, 'spreadsheet.pivot.createDesc')}
        >
          {error && (
            <InlineNotice
              className="work-office-form-error"
              tone="danger"
              role="alert"
            >
              {error}
            </InlineNotice>
          )}
        </StateView>
      )}
    </fieldset>
  );
}

function pivotKey(item: PivotListItem | undefined): string | null {
  return item ? `${item.ownerSheetId}:${item.pivot.id}` : null;
}

function pivotDraftKey(draft: PivotDraft | null): string | null {
  return draft ? `${draft.ownerSheetId}:${draft.id}` : null;
}

function pivotDraft(item: PivotListItem): PivotDraft {
  return {
    ...item.pivot,
    ownerSheetId: item.ownerSheetId,
    rowFields: [...item.pivot.rowFields],
    columnFields: [...item.pivot.columnFields],
    reportFilters:
      item.pivot.reportFilters?.map((filter) => ({ ...filter })) ?? [],
    values: item.pivot.values.map((value) => ({ ...value })),
  };
}

function pivotFieldRole(draft: PivotDraft, fieldIndex: number): PivotFieldRole {
  if (draft.rowFields.includes(fieldIndex)) return 'row';
  if (draft.columnFields.includes(fieldIndex)) return 'column';
  if (draft.reportFilters?.some((filter) => filter.fieldIndex === fieldIndex))
    return 'filter';
  if (draft.values.some((value) => value.fieldIndex === fieldIndex))
    return 'value';
  return 'unused';
}

function assignPivotFieldRole(
  draft: PivotDraft,
  fieldIndex: number,
  role: PivotFieldRole,
  fieldName: string,
  numeric: boolean,
): PivotDraft {
  const next: PivotDraft = {
    ...draft,
    rowFields: draft.rowFields.filter((index) => index !== fieldIndex),
    columnFields: draft.columnFields.filter((index) => index !== fieldIndex),
    reportFilters: (draft.reportFilters ?? []).filter(
      (filter) => filter.fieldIndex !== fieldIndex,
    ),
    values: draft.values.filter((value) => value.fieldIndex !== fieldIndex),
  };
  if (role === 'row') next.rowFields = [...next.rowFields, fieldIndex];
  else if (role === 'column')
    next.columnFields = [...next.columnFields, fieldIndex];
  else if (role === 'filter')
    next.reportFilters = [...(next.reportFilters ?? []), { fieldIndex }];
  else if (role === 'value') {
    const aggregation: WorkSpreadsheetPivotAggregation = numeric
      ? 'sum'
      : 'counta';
    next.values = [
      ...next.values,
      {
        fieldIndex,
        aggregation,
        caption: defaultPivotValueCaption(fieldName, aggregation),
      },
    ];
  }
  return next;
}

function updatePivotValue(
  draft: PivotDraft,
  fieldIndex: number,
  changes: Partial<WorkSpreadsheetPivotTable['values'][number]>,
): PivotDraft {
  return {
    ...draft,
    values: draft.values.map((value) =>
      value.fieldIndex === fieldIndex ? { ...value, ...changes } : value,
    ),
  };
}

function updatePivotReportFilterItems(
  draft: PivotDraft,
  fieldIndex: number,
  selectedItems: WorkSpreadsheetPivotFilterValue[] | undefined,
): PivotDraft {
  return {
    ...draft,
    reportFilters: (draft.reportFilters ?? []).map((filter) =>
      filter.fieldIndex === fieldIndex
        ? {
            fieldIndex,
            ...(selectedItems === undefined ? {} : { selectedItems }),
          }
        : filter,
    ),
  };
}

function togglePivotReportFilterItem(
  draft: PivotDraft,
  fieldIndex: number,
  value: WorkSpreadsheetPivotFilterValue,
  checked: boolean,
  allItems: SpreadsheetPivotFilterItem[],
): PivotDraft {
  const filter = (draft.reportFilters ?? []).find(
    (entry) => entry.fieldIndex === fieldIndex,
  ) ?? { fieldIndex };
  const selection = spreadsheetPivotReportFilterSelection(filter);
  const keys = new Set<string>(
    selection.kind === 'all'
      ? allItems.map((item) => spreadsheetPivotFilterValueKey(item.value))
      : selection.kind === 'items'
        ? selection.items.map(spreadsheetPivotFilterValueKey)
        : [],
  );
  const valueKey = spreadsheetPivotFilterValueKey(value);
  if (checked) keys.add(valueKey);
  else keys.delete(valueKey);
  if (!keys.size) return updatePivotReportFilterItems(draft, fieldIndex, []);
  if (keys.size === allItems.length)
    return updatePivotReportFilterItems(draft, fieldIndex, undefined);
  return updatePivotReportFilterItems(
    draft,
    fieldIndex,
    allItems
      .filter((item) => keys.has(spreadsheetPivotFilterValueKey(item.value)))
      .map((item) => item.value),
  );
}
