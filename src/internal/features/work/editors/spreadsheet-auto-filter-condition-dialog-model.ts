import {
  officeMessage,
  resolveOfficeMessages,
  type OfficeMessageKey,
} from '../../../i18n/office-locale';
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

const CONDITION_LABEL_KEYS: Readonly<
  Record<SpreadsheetAutoFilterConditionType, OfficeMessageKey>
> = {
  equals: 'spreadsheet.autoFilter.op.equals',
  'not-equals': 'spreadsheet.autoFilter.op.not-equals',
  contains: 'spreadsheet.autoFilter.op.contains',
  'does-not-contain': 'spreadsheet.autoFilter.op.does-not-contain',
  'begins-with': 'spreadsheet.autoFilter.op.begins-with',
  'does-not-begin-with': 'spreadsheet.autoFilter.op.does-not-begin-with',
  'ends-with': 'spreadsheet.autoFilter.op.ends-with',
  'does-not-end-with': 'spreadsheet.autoFilter.op.does-not-end-with',
  'matches-wildcard': 'spreadsheet.autoFilter.op.matches-wildcard',
  'does-not-match-wildcard': 'spreadsheet.autoFilter.op.does-not-match-wildcard',
  'greater-than': 'spreadsheet.autoFilter.op.greater-than',
  'greater-than-or-equal': 'spreadsheet.autoFilter.op.greater-than-or-equal',
  'less-than': 'spreadsheet.autoFilter.op.less-than',
  'less-than-or-equal': 'spreadsheet.autoFilter.op.less-than-or-equal',
  between: 'spreadsheet.autoFilter.op.between',
  'not-between': 'spreadsheet.autoFilter.op.not-between',
  top: 'spreadsheet.autoFilter.op.top',
  'top-percent': 'spreadsheet.autoFilter.op.top-percent',
  bottom: 'spreadsheet.autoFilter.op.bottom',
  'bottom-percent': 'spreadsheet.autoFilter.op.bottom-percent',
  'above-average': 'spreadsheet.autoFilter.op.above-average',
  'below-average': 'spreadsheet.autoFilter.op.below-average',
  tomorrow: 'spreadsheet.autoFilter.op.tomorrow',
  today: 'spreadsheet.autoFilter.op.today',
  yesterday: 'spreadsheet.autoFilter.op.yesterday',
  'next-week': 'spreadsheet.autoFilter.op.next-week',
  'this-week': 'spreadsheet.autoFilter.op.this-week',
  'last-week': 'spreadsheet.autoFilter.op.last-week',
  'next-month': 'spreadsheet.autoFilter.op.next-month',
  'this-month': 'spreadsheet.autoFilter.op.this-month',
  'last-month': 'spreadsheet.autoFilter.op.last-month',
  'next-quarter': 'spreadsheet.autoFilter.op.next-quarter',
  'this-quarter': 'spreadsheet.autoFilter.op.this-quarter',
  'last-quarter': 'spreadsheet.autoFilter.op.last-quarter',
  'next-year': 'spreadsheet.autoFilter.op.next-year',
  'this-year': 'spreadsheet.autoFilter.op.this-year',
  'last-year': 'spreadsheet.autoFilter.op.last-year',
  'year-to-date': 'spreadsheet.autoFilter.op.year-to-date',
  'quarter-1': 'spreadsheet.autoFilter.op.quarter-1',
  'quarter-2': 'spreadsheet.autoFilter.op.quarter-2',
  'quarter-3': 'spreadsheet.autoFilter.op.quarter-3',
  'quarter-4': 'spreadsheet.autoFilter.op.quarter-4',
  'month-1': 'spreadsheet.autoFilter.op.month-1',
  'month-2': 'spreadsheet.autoFilter.op.month-2',
  'month-3': 'spreadsheet.autoFilter.op.month-3',
  'month-4': 'spreadsheet.autoFilter.op.month-4',
  'month-5': 'spreadsheet.autoFilter.op.month-5',
  'month-6': 'spreadsheet.autoFilter.op.month-6',
  'month-7': 'spreadsheet.autoFilter.op.month-7',
  'month-8': 'spreadsheet.autoFilter.op.month-8',
  'month-9': 'spreadsheet.autoFilter.op.month-9',
  'month-10': 'spreadsheet.autoFilter.op.month-10',
  'month-11': 'spreadsheet.autoFilter.op.month-11',
  'month-12': 'spreadsheet.autoFilter.op.month-12',
  blanks: 'spreadsheet.autoFilter.op.blanks',
  'non-blanks': 'spreadsheet.autoFilter.op.non-blanks',
};

export function spreadsheetAutoFilterConditionLabel(
  type: SpreadsheetAutoFilterConditionType,
): string {
  return officeMessage(resolveOfficeMessages(), CONDITION_LABEL_KEYS[type]);
}


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
    if (!/^\d+$/.test(value)) return officeMessage(resolveOfficeMessages(), 'spreadsheet.autoFilter.error.integer');
    const maximum =
      draft.type === 'top-percent' || draft.type === 'bottom-percent'
        ? 100
        : 500;
    const numeric = Number(value);
    return numeric >= 1 && numeric <= maximum
      ? null
      : officeMessage(resolveOfficeMessages(), 'spreadsheet.autoFilter.error.integerRange', {
          n: String(maximum),
        });
  }
  const valueError = spreadsheetAutoFilterValueError(draft.type, draft.value);
  if (valueError) return valueError;
  if (NUMBER_CONDITIONS.includes(draft.type)) {
    if (
      (draft.type === 'between' || draft.type === 'not-between') &&
      (!draft.upperValue.trim() || !Number.isFinite(Number(draft.upperValue)))
    ) {
      return officeMessage(resolveOfficeMessages(), 'spreadsheet.autoFilter.error.upper');
    }
    if (
      (draft.type === 'between' || draft.type === 'not-between') &&
      Number(draft.value) > Number(draft.upperValue)
    ) {
      return officeMessage(resolveOfficeMessages(), 'spreadsheet.autoFilter.error.lowerGtUpper');
    }
  }
  return null;
}

export function spreadsheetAutoFilterValueError(
  type: SpreadsheetAutoFilterConditionType,
  value: string,
): string | null {
  if (spreadsheetAutoFilterDynamicConditionType(type)) return null;
  if (!value.trim()) return officeMessage(resolveOfficeMessages(), 'spreadsheet.autoFilter.error.valueRequired');
  if (
    WILDCARD_CONDITIONS.includes(
      type as (typeof WILDCARD_CONDITIONS)[number],
    ) &&
    workSpreadsheetFilterTextCharacters(value) >
      WORK_SPREADSHEET_FILTER_TEXT_MAX_CHARACTERS
  ) {
    return officeMessage(resolveOfficeMessages(), 'spreadsheet.autoFilter.error.wildcardTooLong', {
      n: WORK_SPREADSHEET_FILTER_TEXT_MAX_CHARACTERS.toLocaleString(),
    });
  }
  return NUMBER_CONDITIONS.includes(type) && !Number.isFinite(Number(value))
    ? officeMessage(resolveOfficeMessages(), 'spreadsheet.autoFilter.error.number')
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
