import type {
  WorkSpreadsheetCustomFilterCondition,
  WorkSpreadsheetFilterCriteria,
} from '../work-types';

export type SpreadsheetAutoFilterConditionType =
  | 'equals'
  | 'not-equals'
  | 'contains'
  | 'does-not-contain'
  | 'begins-with'
  | 'does-not-begin-with'
  | 'ends-with'
  | 'does-not-end-with'
  | 'greater-than'
  | 'greater-than-or-equal'
  | 'less-than'
  | 'less-than-or-equal'
  | 'between'
  | 'not-between'
  | 'blanks'
  | 'non-blanks';

export const CONDITION_LABELS: Readonly<
  Record<SpreadsheetAutoFilterConditionType, string>
> = {
  equals: '等于',
  'not-equals': '不等于',
  contains: '包含',
  'does-not-contain': '不包含',
  'begins-with': '开头是',
  'does-not-begin-with': '开头不是',
  'ends-with': '结尾是',
  'does-not-end-with': '结尾不是',
  'greater-than': '大于',
  'greater-than-or-equal': '大于或等于',
  'less-than': '小于',
  'less-than-or-equal': '小于或等于',
  between: '介于',
  'not-between': '不介于',
  blanks: '空白',
  'non-blanks': '非空白',
};

export const TEXT_CONDITIONS: readonly SpreadsheetAutoFilterConditionType[] = [
  'equals',
  'not-equals',
  'contains',
  'does-not-contain',
  'begins-with',
  'does-not-begin-with',
  'ends-with',
  'does-not-end-with',
];

export const NUMBER_COMPARISON_CONDITIONS = [
  'greater-than',
  'greater-than-or-equal',
  'less-than',
  'less-than-or-equal',
] as const satisfies readonly SpreadsheetAutoFilterConditionType[];

export const NUMBER_CONDITIONS: readonly SpreadsheetAutoFilterConditionType[] =
  [...NUMBER_COMPARISON_CONDITIONS, 'between', 'not-between'];

export const BLANK_CONDITIONS: readonly SpreadsheetAutoFilterConditionType[] = [
  'blanks',
  'non-blanks',
];

export interface SpreadsheetAutoFilterConditionDraft {
  conjunction: 'and' | 'or';
  secondType: WorkSpreadsheetCustomFilterCondition['type'];
  secondValue: string;
  type: SpreadsheetAutoFilterConditionType;
  upperValue: string;
  useSecond: boolean;
  value: string;
}

export function spreadsheetAutoFilterConditionDraft(
  criteria: WorkSpreadsheetFilterCriteria | null,
  numeric: boolean,
): SpreadsheetAutoFilterConditionDraft {
  if (criteria?.type === 'compound') {
    return {
      conjunction: criteria.conjunction,
      secondType: criteria.conditions[1].type,
      secondValue: criteria.conditions[1].value,
      type: criteria.conditions[0].type,
      upperValue: '',
      useSecond: true,
      value: criteria.conditions[0].value,
    };
  }
  const defaults = {
    conjunction: 'and' as const,
    secondType: (numeric ? 'greater-than' : 'equals') as
      | 'greater-than'
      | 'equals',
    secondValue: '',
    useSecond: false,
  };
  if (criteria && spreadsheetAutoFilterConditionType(criteria.type)) {
    if (criteria.type === 'between' || criteria.type === 'not-between') {
      return {
        ...defaults,
        type: criteria.type,
        value: criteria.lower,
        upperValue: criteria.upper,
      };
    }
    if (criteria.type === 'blanks' || criteria.type === 'non-blanks') {
      return {
        ...defaults,
        type: criteria.type,
        value: '',
        upperValue: '',
      };
    }
    if ('value' in criteria) {
      return {
        ...defaults,
        type: criteria.type,
        value: criteria.value,
        upperValue: '',
      };
    }
  }
  return {
    ...defaults,
    type: numeric ? 'equals' : 'contains',
    value: '',
    upperValue: '',
  };
}

export function spreadsheetAutoFilterConditionCriteria(
  draft: SpreadsheetAutoFilterConditionDraft,
): WorkSpreadsheetFilterCriteria | null {
  if (
    spreadsheetAutoFilterPrimaryConditionError(draft) ||
    (draft.useSecond &&
      spreadsheetAutoFilterValueError(draft.secondType, draft.secondValue))
  ) {
    return null;
  }
  if (draft.useSecond) {
    if (!spreadsheetAutoFilterCustomConditionType(draft.type)) return null;
    return {
      type: 'compound',
      conjunction: draft.conjunction,
      conditions: [
        spreadsheetAutoFilterCustomCondition(draft.type, draft.value),
        spreadsheetAutoFilterCustomCondition(
          draft.secondType,
          draft.secondValue,
        ),
      ],
    };
  }
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
  return spreadsheetAutoFilterCustomCondition(draft.type, draft.value);
}

export function spreadsheetAutoFilterPrimaryConditionError(
  draft: SpreadsheetAutoFilterConditionDraft,
): string | null {
  if (draft.type === 'blanks' || draft.type === 'non-blanks') return null;
  const valueError = spreadsheetAutoFilterValueError(draft.type, draft.value);
  if (valueError) return valueError;
  if (NUMBER_CONDITIONS.includes(draft.type)) {
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

export function spreadsheetAutoFilterValueError(
  type: SpreadsheetAutoFilterConditionType,
  value: string,
): string | null {
  if (!value.trim()) return '请输入筛选值。';
  return NUMBER_CONDITIONS.includes(type) && !Number.isFinite(Number(value))
    ? '请输入有效数字。'
    : null;
}

export function spreadsheetAutoFilterCustomConditionType(
  value: string,
): value is WorkSpreadsheetCustomFilterCondition['type'] {
  return (
    TEXT_CONDITIONS.includes(value as SpreadsheetAutoFilterConditionType) ||
    NUMBER_COMPARISON_CONDITIONS.includes(
      value as (typeof NUMBER_COMPARISON_CONDITIONS)[number],
    )
  );
}

function spreadsheetAutoFilterCustomCondition(
  type: WorkSpreadsheetCustomFilterCondition['type'],
  value: string,
): WorkSpreadsheetCustomFilterCondition {
  return {
    type,
    value: NUMBER_CONDITIONS.includes(type) ? value.trim() : value,
  } as WorkSpreadsheetCustomFilterCondition;
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
