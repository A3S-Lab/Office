import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import type {
  WorkSpreadsheetCustomFilterCondition,
  WorkSpreadsheetFilterCriteria,
} from '../work-types';
import {
  BLANK_CONDITIONS,
  CONDITION_LABELS,
  NUMBER_COMPARISON_CONDITIONS,
  NUMBER_CONDITIONS,
  spreadsheetAutoFilterConditionCriteria,
  spreadsheetAutoFilterConditionDraft,
  spreadsheetAutoFilterCustomConditionType,
  spreadsheetAutoFilterPrimaryConditionError,
  spreadsheetAutoFilterValueError,
  TEXT_CONDITIONS,
  type SpreadsheetAutoFilterConditionType,
} from './spreadsheet-auto-filter-condition-dialog-model';

export type { SpreadsheetAutoFilterConditionType } from './spreadsheet-auto-filter-condition-dialog-model';

export interface SpreadsheetAutoFilterConditionDialogSource {
  columnLabel: string;
  criteria: WorkSpreadsheetFilterCriteria | null;
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
  const [draft, setDraft] = useState(() =>
    spreadsheetAutoFilterConditionDraft(source.criteria, source.numeric),
  );
  const [touched, setTouched] = useState(false);
  const formId = useId();
  const primaryError = spreadsheetAutoFilterPrimaryConditionError(draft);
  const secondError = draft.useSecond
    ? spreadsheetAutoFilterValueError(draft.secondType, draft.secondValue)
    : null;
  const error = primaryError ?? secondError;
  const needsValue = !BLANK_CONDITIONS.includes(draft.type);
  const needsUpperValue =
    draft.type === 'between' || draft.type === 'not-between';
  const numericValue = NUMBER_CONDITIONS.includes(draft.type);
  const canUseSecond = spreadsheetAutoFilterCustomConditionType(draft.type);
  const secondNumericValue = NUMBER_CONDITIONS.includes(draft.secondType);
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setTouched(true);
    const criteria = spreadsheetAutoFilterConditionCriteria(draft);
    if (criteria && onApply(criteria)) onClose();
  };

  return (
    <Dialog
      title="自定义自动筛选"
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
              清除此列筛选
            </Button>
          )}
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={Boolean(error)}
          >
            确定
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
          label="筛选条件"
          description="仅隐藏不符合条件的行；其他列的筛选条件会继续生效。"
        >
          <select
            aria-label="筛选条件"
            data-autofocus
            value={draft.type}
            onChange={(event) => {
              const type = event.currentTarget
                .value as SpreadsheetAutoFilterConditionType;
              setDraft((current) => ({
                ...current,
                type,
                useSecond:
                  current.useSecond &&
                  spreadsheetAutoFilterCustomConditionType(type),
                ...(BLANK_CONDITIONS.includes(type)
                  ? { value: '', upperValue: '' }
                  : !NUMBER_CONDITIONS.includes(type)
                    ? { upperValue: '' }
                    : {}),
              }));
              setTouched(false);
            }}
          >
            <optgroup label="文本与值">
              {TEXT_CONDITIONS.map((type) => (
                <option key={type} value={type}>
                  {CONDITION_LABELS[type]}
                </option>
              ))}
            </optgroup>
            <optgroup label="数字">
              {NUMBER_CONDITIONS.map((type) => (
                <option key={type} value={type}>
                  {CONDITION_LABELS[type]}
                </option>
              ))}
            </optgroup>
            <optgroup label="空白单元格">
              {BLANK_CONDITIONS.map((type) => (
                <option key={type} value={type}>
                  {CONDITION_LABELS[type]}
                </option>
              ))}
            </optgroup>
          </select>
        </Field>

        {needsValue && (
          <div className="work-spreadsheet-auto-filter-values">
            <Field
              label={needsUpperValue ? '下限' : '筛选值'}
              required
              error={touched ? (primaryError ?? undefined) : undefined}
            >
              <input
                type="text"
                aria-label={needsUpperValue ? '下限' : '筛选值'}
                inputMode={numericValue ? 'decimal' : 'text'}
                value={draft.value}
                onBlur={() => setTouched(true)}
                onChange={(event) => {
                  const value = event.currentTarget.value;
                  setDraft((current) => ({ ...current, value }));
                }}
              />
            </Field>
            {needsUpperValue && (
              <Field label="上限" required>
                <input
                  type="text"
                  aria-label="上限"
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
            添加第二个条件
          </Button>
        )}

        {draft.useSecond && (
          <div className="work-spreadsheet-auto-filter-compound">
            <fieldset className="work-spreadsheet-auto-filter-conjunction">
              <legend>条件关系</legend>
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
                并且
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
                或者
              </label>
            </fieldset>
            <div className="work-spreadsheet-auto-filter-second-condition">
              <Field label="第二个筛选条件">
                <select
                  aria-label="第二个筛选条件"
                  value={draft.secondType}
                  onChange={(event) => {
                    const secondType = event.currentTarget
                      .value as WorkSpreadsheetCustomFilterCondition['type'];
                    setDraft((current) => ({ ...current, secondType }));
                    setTouched(false);
                  }}
                >
                  <optgroup label="文本与值">
                    {TEXT_CONDITIONS.map((type) => (
                      <option key={type} value={type}>
                        {CONDITION_LABELS[type]}
                      </option>
                    ))}
                  </optgroup>
                  <optgroup label="数字">
                    {NUMBER_COMPARISON_CONDITIONS.map((type) => (
                      <option key={type} value={type}>
                        {CONDITION_LABELS[type]}
                      </option>
                    ))}
                  </optgroup>
                </select>
              </Field>
              <Field
                label="第二个筛选值"
                required
                error={touched ? (secondError ?? undefined) : undefined}
              >
                <input
                  type="text"
                  aria-label="第二个筛选值"
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
              移除第二个条件
            </Button>
          </div>
        )}
      </form>
    </Dialog>
  );
}
