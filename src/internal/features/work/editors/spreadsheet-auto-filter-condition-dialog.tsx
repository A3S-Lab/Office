import { type FormEvent, useId, useMemo, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import type {
  WorkSpreadsheetCustomFilterCondition,
  WorkSpreadsheetFilterCriteria,
} from '../work-types';
import { OfficeSelect, type OfficeSelectOption } from './office-controls';
import {
  AVERAGE_CONDITIONS,
  BLANK_CONDITIONS,
  spreadsheetAutoFilterConditionLabel,
  DATE_CONDITIONS,
  NUMBER_COMPARISON_CONDITIONS,
  NUMBER_CONDITIONS,
  RANK_CONDITIONS,
  spreadsheetAutoFilterConditionCriteria,
  spreadsheetAutoFilterConditionDraft,
  spreadsheetAutoFilterCustomConditionType,
  spreadsheetAutoFilterDynamicConditionType,
  spreadsheetAutoFilterPrimaryConditionError,
  spreadsheetAutoFilterValueError,
  TEXT_CONDITIONS,
  type SpreadsheetAutoFilterConditionType,
  WILDCARD_CONDITIONS,
} from './spreadsheet-auto-filter-condition-dialog-model';
import { useOfficeMessages } from './office-messages-context';

export type { SpreadsheetAutoFilterConditionType } from './spreadsheet-auto-filter-condition-dialog-model';

export interface SpreadsheetAutoFilterConditionDialogSource {
  columnLabel: string;
  criteria: WorkSpreadsheetFilterCriteria | null;
  date: boolean;
  hasActiveFilter: boolean;
  numeric: boolean;
  sheetName: string;
}

export function SpreadsheetAutoFilterConditionDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClear,
  onClose,
}: {
  source: SpreadsheetAutoFilterConditionDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (criteria: WorkSpreadsheetFilterCriteria) => boolean;
  onClear: () => boolean;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const [draft, setDraft] = useState(() =>
    spreadsheetAutoFilterConditionDraft(
      source.criteria,
      source.numeric,
      source.date,
    ),
  );
  const [touched, setTouched] = useState(false);
  const formId = useId();
  const primaryError = spreadsheetAutoFilterPrimaryConditionError(draft);
  const secondError = draft.useSecond
    ? spreadsheetAutoFilterValueError(draft.secondType, draft.secondValue)
    : null;
  const error = primaryError ?? secondError;
  const dynamicCondition = spreadsheetAutoFilterDynamicConditionType(
    draft.type,
  );
  const needsValue =
    !BLANK_CONDITIONS.includes(draft.type) && !dynamicCondition;
  const needsUpperValue =
    draft.type === 'between' || draft.type === 'not-between';
  const rankValue = RANK_CONDITIONS.includes(
    draft.type as (typeof RANK_CONDITIONS)[number],
  );
  const rankPercent =
    draft.type === 'top-percent' || draft.type === 'bottom-percent';
  const wildcardValue = WILDCARD_CONDITIONS.includes(
    draft.type as (typeof WILDCARD_CONDITIONS)[number],
  );
  const valueLabel = rankValue
    ? rankPercent
      ? officeMessage(messages, 'spreadsheet.autoFilter.rank.percent')
      : officeMessage(messages, 'spreadsheet.autoFilter.rank.items')
    : wildcardValue
      ? officeMessage(messages, 'spreadsheet.autoFilter.field.wildcard')
      : needsUpperValue
        ? officeMessage(messages, 'spreadsheet.autoFilter.field.lower')
        : officeMessage(messages, 'spreadsheet.autoFilter.field.value');
  const numericValue = NUMBER_CONDITIONS.includes(draft.type) || rankValue;
  const showRankConditions = source.numeric || rankValue;
  const showAverageConditions =
    source.numeric ||
    AVERAGE_CONDITIONS.includes(
      draft.type as (typeof AVERAGE_CONDITIONS)[number],
    );
  const showDateConditions =
    source.date ||
    DATE_CONDITIONS.includes(draft.type as (typeof DATE_CONDITIONS)[number]);
  const canUseSecond = spreadsheetAutoFilterCustomConditionType(draft.type);
  const secondNumericValue = NUMBER_CONDITIONS.includes(draft.secondType);
  const secondWildcardValue = WILDCARD_CONDITIONS.includes(
    draft.secondType as (typeof WILDCARD_CONDITIONS)[number],
  );
  const primaryConditionOptions = useMemo(() => {
    const options: OfficeSelectOption<SpreadsheetAutoFilterConditionType>[] =
      [];
    for (const type of TEXT_CONDITIONS) {
      options.push({
        value: type,
        label: spreadsheetAutoFilterConditionLabel(type),
        group: officeMessage(messages, 'spreadsheet.autoFilter.group.textValue'),
      });
    }
    for (const type of NUMBER_CONDITIONS) {
      options.push({
        value: type,
        label: spreadsheetAutoFilterConditionLabel(type),
        group: officeMessage(messages, 'spreadsheet.autoFilter.group.number'),
      });
    }
    if (showRankConditions) {
      for (const type of RANK_CONDITIONS) {
        options.push({
          value: type,
          label: spreadsheetAutoFilterConditionLabel(type),
          group: officeMessage(messages, 'spreadsheet.autoFilter.group.rank'),
        });
      }
    }
    if (showAverageConditions) {
      for (const type of AVERAGE_CONDITIONS) {
        options.push({
          value: type,
          label: spreadsheetAutoFilterConditionLabel(type),
          group: officeMessage(messages, 'spreadsheet.autoFilter.group.average'),
        });
      }
    }
    if (showDateConditions) {
      for (const type of DATE_CONDITIONS) {
        options.push({
          value: type,
          label: spreadsheetAutoFilterConditionLabel(type),
          group: officeMessage(messages, 'spreadsheet.autoFilter.group.date'),
        });
      }
    }
    for (const type of BLANK_CONDITIONS) {
      options.push({
        value: type,
        label: spreadsheetAutoFilterConditionLabel(type),
        group: officeMessage(messages, 'spreadsheet.autoFilter.group.blanks'),
      });
    }
    return options;
  }, [showAverageConditions, showDateConditions, showRankConditions]);
  const secondConditionOptions = useMemo(() => {
    const options: OfficeSelectOption<
      WorkSpreadsheetCustomFilterCondition['type']
    >[] = [];
    for (const type of TEXT_CONDITIONS) {
      options.push({
        value: type as WorkSpreadsheetCustomFilterCondition['type'],
        label: spreadsheetAutoFilterConditionLabel(type),
        group: officeMessage(messages, 'spreadsheet.autoFilter.group.textValue'),
      });
    }
    for (const type of NUMBER_COMPARISON_CONDITIONS) {
      options.push({
        value: type as WorkSpreadsheetCustomFilterCondition['type'],
        label: spreadsheetAutoFilterConditionLabel(type),
        group: officeMessage(messages, 'spreadsheet.autoFilter.group.number'),
      });
    }
    return options;
  }, []);
  const changePrimaryType = (type: SpreadsheetAutoFilterConditionType) => {
    setDraft((current) => ({
      ...current,
      type,
      useSecond:
        current.useSecond && spreadsheetAutoFilterCustomConditionType(type),
      ...(BLANK_CONDITIONS.includes(type) ||
      spreadsheetAutoFilterDynamicConditionType(type)
        ? { value: '', upperValue: '' }
        : !NUMBER_CONDITIONS.includes(type)
          ? { upperValue: '' }
          : {}),
    }));
    setTouched(false);
  };
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setTouched(true);
    const criteria = spreadsheetAutoFilterConditionCriteria(draft);
    if (criteria && onApply(criteria)) onClose();
  };

  return (
    <Dialog
      title={officeMessage(messages, 'spreadsheet.autoFilter.title')}
      description={`${source.sheetName}!${source.columnLabel}`}
      className="work-spreadsheet-auto-filter-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          {source.hasActiveFilter && (
            <Button
              tone="danger"
              className="work-spreadsheet-auto-filter-clear"
              onClick={() => {
                if (onClear()) onClose();
              }}
            >
              {officeMessage(messages, 'spreadsheet.autoFilter.clearColumn')}
            </Button>
          )}
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.autoFilter.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={Boolean(error)}
          >
            {officeMessage(messages, 'spreadsheet.autoFilter.ok')}
          </Button>
        </>
      }
    >
      <form
        id={formId}
        className="work-spreadsheet-auto-filter-form"
        onSubmit={submit}
      >
        <Field
          label={officeMessage(messages, 'spreadsheet.autoFilter.condition')}
          description={officeMessage(messages, 'spreadsheet.autoFilter.conditionDesc')}
        >
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.autoFilter.condition')}
            initialFocus
            value={draft.type}
            options={primaryConditionOptions}
            onValueChange={changePrimaryType}
          />
        </Field>

        {needsValue && (
          <div className="work-spreadsheet-auto-filter-values">
            <Field
              label={valueLabel}
              required
              description={
                wildcardValue
                  ? officeMessage(messages, 'spreadsheet.autoFilter.wildcardHint')
                  : undefined
              }
              error={touched ? (primaryError ?? undefined) : undefined}
            >
              <input
                type="text"
                aria-label={valueLabel}
                inputMode={
                  rankValue ? 'numeric' : numericValue ? 'decimal' : 'text'
                }
                value={draft.value}
                onBlur={() => setTouched(true)}
                onChange={(event) => {
                  const value = event.currentTarget.value;
                  setDraft((current) => ({ ...current, value }));
                }}
              />
            </Field>
            {needsUpperValue && (
              <Field label={officeMessage(messages, 'spreadsheet.autoFilter.upper')} required>
                <input
                  type="text"
                  aria-label={officeMessage(messages, 'spreadsheet.autoFilter.upper')}
                  inputMode="decimal"
                  value={draft.upperValue}
                  onBlur={() => setTouched(true)}
                  onChange={(event) => {
                    const upperValue = event.currentTarget.value;
                    setDraft((current) => ({ ...current, upperValue }));
                  }}
                />
              </Field>
            )}
          </div>
        )}

        {canUseSecond && !draft.useSecond && (
          <Button
            type="button"
            tone="quiet"
            className="work-spreadsheet-auto-filter-add-condition"
            onClick={() => {
              setDraft((current) => ({ ...current, useSecond: true }));
              setTouched(false);
            }}
          >
            {officeMessage(messages, 'spreadsheet.autoFilter.addSecond')}
          </Button>
        )}

        {draft.useSecond && (
          <div className="work-spreadsheet-auto-filter-compound">
            <fieldset className="work-spreadsheet-auto-filter-conjunction">
              <legend>{officeMessage(messages, 'spreadsheet.autoFilter.relation')}</legend>
              <label>
                <input
                  type="radio"
                  name={`${formId}-conjunction`}
                  checked={draft.conjunction === 'and'}
                  onChange={() => {
                    setDraft((current) => ({
                      ...current,
                      conjunction: 'and',
                    }));
                  }}
                />
                {officeMessage(messages, 'spreadsheet.autoFilter.and')}
              </label>
              <label>
                <input
                  type="radio"
                  name={`${formId}-conjunction`}
                  checked={draft.conjunction === 'or'}
                  onChange={() => {
                    setDraft((current) => ({
                      ...current,
                      conjunction: 'or',
                    }));
                  }}
                />
                {officeMessage(messages, 'spreadsheet.autoFilter.or')}
              </label>
            </fieldset>
            <div className="work-spreadsheet-auto-filter-second-condition">
              <Field label={officeMessage(messages, 'spreadsheet.autoFilter.secondCondition')}>
                <OfficeSelect
                  ariaLabel={officeMessage(messages, 'spreadsheet.autoFilter.secondCondition')}
                  value={draft.secondType}
                  options={secondConditionOptions}
                  onValueChange={(secondType) => {
                    setDraft((current) => ({ ...current, secondType }));
                    setTouched(false);
                  }}
                />
              </Field>
              <Field
                label={officeMessage(messages, 'spreadsheet.autoFilter.secondValue')}
                required
                description={
                  secondWildcardValue
                    ? officeMessage(messages, 'spreadsheet.autoFilter.wildcardHint')
                    : undefined
                }
                error={touched ? (secondError ?? undefined) : undefined}
              >
                <input
                  type="text"
                  aria-label={officeMessage(messages, 'spreadsheet.autoFilter.secondValue')}
                  inputMode={secondNumericValue ? 'decimal' : 'text'}
                  value={draft.secondValue}
                  onBlur={() => setTouched(true)}
                  onChange={(event) => {
                    const secondValue = event.currentTarget.value;
                    setDraft((current) => ({ ...current, secondValue }));
                  }}
                />
              </Field>
            </div>
            <Button
              type="button"
              tone="quiet"
              className="work-spreadsheet-auto-filter-remove-condition"
              onClick={() => {
                setDraft((current) => ({
                  ...current,
                  secondValue: '',
                  useSecond: false,
                }));
                setTouched(false);
              }}
            >
              {officeMessage(messages, 'spreadsheet.autoFilter.removeSecond')}
            </Button>
          </div>
        )}
      </form>
    </Dialog>
  );
}
