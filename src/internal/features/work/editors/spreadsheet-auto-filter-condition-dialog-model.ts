import type {
  WorkSpreadsheetCustomFilterCondition,
  WorkSpreadsheetDynamicFilter,
  WorkSpreadsheetFilterCriteria,
} from '../work-types';
import {
  WORK_SPREADSHEET_FILTER_TEXT_MAX_CHARACTERS,
  workSpreadsheetFilterTextCharacters,
} from '../work-spreadsheet-filter-contract';

export type SpreadsheetAutoFilterConditionType =
  | WorkSpreadsheetCustomFilterCondition['type']
  | WorkSpreadsheetDynamicFilter
  | 'between'
  | 'not-between'
  | 'top'
  | 'top-percent'
  | 'bottom'
  | 'bottom-percent'
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
  'matches-wildcard': '通配符匹配',
  'does-not-match-wildcard': '通配符不匹配',
  'greater-than': '大于',
  'greater-than-or-equal': '大于或等于',
  'less-than': '小于',
  'less-than-or-equal': '小于或等于',
  between: '介于',
  'not-between': '不介于',
  top: '前几项',
  'top-percent': '前百分比',
  bottom: '后几项',
  'bottom-percent': '后百分比',
  'above-average': '高于平均值',
  'below-average': '低于平均值',
  tomorrow: '明天',
  today: '今天',
  yesterday: '昨天',
  'next-week': '下周',
  'this-week': '本周',
  'last-week': '上周',
  'next-month': '下月',
  'this-month': '本月',
  'last-month': '上月',
  'next-quarter': '下季度',
  'this-quarter': '本季度',
  'last-quarter': '上季度',
  'next-year': '明年',
  'this-year': '今年',
  'last-year': '去年',
  'year-to-date': '年初至今',
  'quarter-1': '第一季度',
  'quarter-2': '第二季度',
  'quarter-3': '第三季度',
  'quarter-4': '第四季度',
  'month-1': '一月',
  'month-2': '二月',
  'month-3': '三月',
  'month-4': '四月',
  'month-5': '五月',
  'month-6': '六月',
  'month-7': '七月',
  'month-8': '八月',
  'month-9': '九月',
  'month-10': '十月',
  'month-11': '十一月',
  'month-12': '十二月',
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
  'matches-wildcard',
  'does-not-match-wildcard',
];

export const WILDCARD_CONDITIONS = [
  'matches-wildcard',
  'does-not-match-wildcard',
] as const satisfies readonly SpreadsheetAutoFilterConditionType[];

export const NUMBER_COMPARISON_CONDITIONS = [
  'greater-than',
  'greater-than-or-equal',
  'less-than',
  'less-than-or-equal',
] as const satisfies readonly SpreadsheetAutoFilterConditionType[];

export const NUMBER_CONDITIONS: readonly SpreadsheetAutoFilterConditionType[] =
  [...NUMBER_COMPARISON_CONDITIONS, 'between', 'not-between'];

export const RANK_CONDITIONS = [
  'top',
  'top-percent',
  'bottom',
  'bottom-percent',
] as const satisfies readonly SpreadsheetAutoFilterConditionType[];

export const AVERAGE_CONDITIONS = [
  'above-average',
  'below-average',
] as const satisfies readonly WorkSpreadsheetDynamicFilter[];

export const DATE_CONDITIONS = [
  'today',
  'yesterday',
  'tomorrow',
  'this-week',
  'last-week',
  'next-week',
  'this-month',
  'last-month',
  'next-month',
  'this-quarter',
  'last-quarter',
  'next-quarter',
  'this-year',
  'last-year',
  'next-year',
  'year-to-date',
  'quarter-1',
  'quarter-2',
  'quarter-3',
  'quarter-4',
  'month-1',
  'month-2',
  'month-3',
  'month-4',
  'month-5',
  'month-6',
  'month-7',
  'month-8',
  'month-9',
  'month-10',
  'month-11',
  'month-12',
] as const satisfies readonly WorkSpreadsheetDynamicFilter[];

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
  date: boolean,
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
  if (
    criteria?.type === 'top' ||
    criteria?.type === 'bottom' ||
    criteria?.type === 'top-percent' ||
    criteria?.type === 'bottom-percent'
  ) {
    const value =
      criteria.type === 'top' || criteria.type === 'bottom'
        ? criteria.count
        : criteria.percent;
    return {
      conjunction: 'and',
      secondType: 'greater-than',
      secondValue: '',
      type: criteria.type,
      upperValue: '',
      useSecond: false,
      value: String(value),
    };
  }
  if (criteria?.type === 'dynamic') {
    return {
      conjunction: 'and',
      secondType: numeric ? 'greater-than' : 'equals',
      secondValue: '',
      type: criteria.kind,
      upperValue: '',
      useSecond: false,
      value: '',
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
    type: date ? 'today' : numeric ? 'equals' : 'contains',
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
  if (spreadsheetAutoFilterDynamicConditionType(draft.type)) {
    return { type: 'dynamic', kind: draft.type };
  }
  if (spreadsheetAutoFilterRankConditionType(draft.type)) {
    const value = Number(draft.value.trim());
    return draft.type === 'top' || draft.type === 'bottom'
      ? { type: draft.type, count: value }
      : { type: draft.type, percent: value };
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
  if (
    draft.type === 'blanks' ||
    draft.type === 'non-blanks' ||
    spreadsheetAutoFilterDynamicConditionType(draft.type)
  ) {
    return null;
  }
  if (spreadsheetAutoFilterRankConditionType(draft.type)) {
    const value = draft.value.trim();
    if (!/^\d+$/.test(value)) return '请输入整数。';
    const maximum =
      draft.type === 'top-percent' || draft.type === 'bottom-percent'
        ? 100
        : 500;
    const numeric = Number(value);
    return numeric >= 1 && numeric <= maximum
      ? null
      : `请输入 1 到 ${maximum} 之间的整数。`;
  }
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
  if (spreadsheetAutoFilterDynamicConditionType(type)) return null;
  if (!value.trim()) return '请输入筛选值。';
  if (
    WILDCARD_CONDITIONS.includes(
      type as (typeof WILDCARD_CONDITIONS)[number],
    ) &&
    workSpreadsheetFilterTextCharacters(value) >
      WORK_SPREADSHEET_FILTER_TEXT_MAX_CHARACTERS
  ) {
    return `通配符表达式不能超过 ${WORK_SPREADSHEET_FILTER_TEXT_MAX_CHARACTERS.toLocaleString()} 个字符。`;
  }
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

export function spreadsheetAutoFilterRankConditionType(
  value: string,
): value is (typeof RANK_CONDITIONS)[number] {
  return RANK_CONDITIONS.includes(value as (typeof RANK_CONDITIONS)[number]);
}

export function spreadsheetAutoFilterDynamicConditionType(
  value: string,
): value is WorkSpreadsheetDynamicFilter {
  return (
    AVERAGE_CONDITIONS.includes(value as (typeof AVERAGE_CONDITIONS)[number]) ||
    DATE_CONDITIONS.includes(value as (typeof DATE_CONDITIONS)[number])
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
    RANK_CONDITIONS.includes(value as (typeof RANK_CONDITIONS)[number]) ||
    BLANK_CONDITIONS.includes(value as SpreadsheetAutoFilterConditionType) ||
    spreadsheetAutoFilterDynamicConditionType(value)
  );
}
