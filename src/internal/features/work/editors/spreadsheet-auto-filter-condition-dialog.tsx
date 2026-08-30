import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import type { WorkSpreadsheetFilterCriteria } from '../work-types';

export type SpreadsheetAutoFilterConditionType =
  | 'equals'
  | 'not-equals'
  | 'contains'
  | 'does-not-contain'
  | 'begins-with'
  | 'ends-with'
  | 'greater-than'
  | 'greater-than-or-equal'
  | 'less-than'
  | 'less-than-or-equal'
  | 'between'
  | 'not-between'
  | 'blanks'
  | 'non-blanks';

export interface SpreadsheetAutoFilterConditionDialogSource {
  columnLabel: string;
  criteria: WorkSpreadsheetFilterCriteria | null;
  hasActiveFilter: boolean;
  numeric: boolean;
  sheetName: string;
}

const CONDITION_LABELS: Readonly<
  Record<SpreadsheetAutoFilterConditionType, string>
> = {
  equals: '等于',
  'not-equals': '不等于',
  contains: '包含',
  'does-not-contain': '不包含',
  'begins-with': '开头是',
  'ends-with': '结尾是',
  'greater-than': '大于',
  'greater-than-or-equal': '大于或等于',
  'less-than': '小于',
  'less-than-or-equal': '小于或等于',
  between: '介于',
  'not-between': '不介于',
  blanks: '空白',
  'non-blanks': '非空白',
};

const TEXT_CONDITIONS: readonly SpreadsheetAutoFilterConditionType[] = [
  'equals',
  'not-equals',
  'contains',
  'does-not-contain',
  'begins-with',
  'ends-with',
];

const NUMBER_CONDITIONS: readonly SpreadsheetAutoFilterConditionType[] = [
  'greater-than',
  'greater-than-or-equal',
  'less-than',
  'less-than-or-equal',
  'between',
  'not-between',
];

const BLANK_CONDITIONS: readonly SpreadsheetAutoFilterConditionType[] = [
  'blanks',
  'non-blanks',
];

interface SpreadsheetAutoFilterConditionDraft {
  type: SpreadsheetAutoFilterConditionType;
  upperValue: string;
  value: string;
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
  const error = spreadsheetAutoFilterConditionError(draft);
  const needsValue = !BLANK_CONDITIONS.includes(draft.type);
  const needsUpperValue =
    draft.type === 'between' || draft.type === 'not-between';
  const numericValue = NUMBER_CONDITIONS.includes(draft.type);
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
              error={touched ? (error ?? undefined) : undefined}
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
      </form>
    </Dialog>
  );
}

function spreadsheetAutoFilterConditionDraft(
  criteria: WorkSpreadsheetFilterCriteria | null,
  numeric: boolean,
): SpreadsheetAutoFilterConditionDraft {
  if (criteria && spreadsheetAutoFilterConditionType(criteria.type)) {
    if (criteria.type === 'between' || criteria.type === 'not-between') {
      return {
        type: criteria.type,
        value: criteria.lower,
        upperValue: criteria.upper,
      };
    }
    if (criteria.type === 'blanks' || criteria.type === 'non-blanks') {
      return { type: criteria.type, value: '', upperValue: '' };
    }
    if ('value' in criteria) {
      return { type: criteria.type, value: criteria.value, upperValue: '' };
    }
  }
  return {
    type: numeric ? 'equals' : 'contains',
    value: '',
    upperValue: '',
  };
}

function spreadsheetAutoFilterConditionCriteria(
  draft: SpreadsheetAutoFilterConditionDraft,
): WorkSpreadsheetFilterCriteria | null {
  if (spreadsheetAutoFilterConditionError(draft)) return null;
  if (draft.type === 'blanks' || draft.type === 'non-blanks') {
    return { type: draft.type };
  }
  if (draft.type === 'between' || draft.type === 'not-between') {
    return {
      type: draft.type,
      lower: draft.value.trim(),
      upper: draft.upperValue.trim(),
    };
  }
  return { type: draft.type, value: draft.value };
}

function spreadsheetAutoFilterConditionError(
  draft: SpreadsheetAutoFilterConditionDraft,
): string | null {
  if (draft.type === 'blanks' || draft.type === 'non-blanks') return null;
  if (!draft.value.trim()) return '请输入筛选值。';
  if (NUMBER_CONDITIONS.includes(draft.type)) {
    if (!Number.isFinite(Number(draft.value))) return '请输入有效数字。';
    if (
      (draft.type === 'between' || draft.type === 'not-between') &&
      (!draft.upperValue.trim() || !Number.isFinite(Number(draft.upperValue)))
    ) {
      return '请输入有效的上限。';
    }
    if (
      (draft.type === 'between' || draft.type === 'not-between') &&
      Number(draft.value) > Number(draft.upperValue)
    ) {
      return '下限不能大于上限。';
    }
  }
  return null;
}

function spreadsheetAutoFilterConditionType(
  value: string,
): value is SpreadsheetAutoFilterConditionType {
  return (
    TEXT_CONDITIONS.includes(value as SpreadsheetAutoFilterConditionType) ||
    NUMBER_CONDITIONS.includes(value as SpreadsheetAutoFilterConditionType) ||
    BLANK_CONDITIONS.includes(value as SpreadsheetAutoFilterConditionType)
  );
}
