import { attribute, directChild, directChildren } from './work-ooxml-package';
import type {
  WorkSpreadsheetDynamicFilter,
  WorkSpreadsheetFilter,
  WorkSpreadsheetFilterCriteria,
  WorkSpreadsheetTableFilter,
} from './work-types';

const DYNAMIC_FILTERS = new Map<string, WorkSpreadsheetDynamicFilter>([
  ['aboveAverage', 'above-average'],
  ['belowAverage', 'below-average'],
  ['tomorrow', 'tomorrow'],
  ['today', 'today'],
  ['yesterday', 'yesterday'],
  ['nextWeek', 'next-week'],
  ['thisWeek', 'this-week'],
  ['lastWeek', 'last-week'],
  ['nextMonth', 'next-month'],
  ['thisMonth', 'this-month'],
  ['lastMonth', 'last-month'],
  ['nextQuarter', 'next-quarter'],
  ['thisQuarter', 'this-quarter'],
  ['lastQuarter', 'last-quarter'],
  ['nextYear', 'next-year'],
  ['thisYear', 'this-year'],
  ['lastYear', 'last-year'],
  ['yearToDate', 'year-to-date'],
  ['Q1', 'quarter-1'],
  ['Q2', 'quarter-2'],
  ['Q3', 'quarter-3'],
  ['Q4', 'quarter-4'],
  ['M1', 'month-1'],
  ['M2', 'month-2'],
  ['M3', 'month-3'],
  ['M4', 'month-4'],
  ['M5', 'month-5'],
  ['M6', 'month-6'],
  ['M7', 'month-7'],
  ['M8', 'month-8'],
  ['M9', 'month-9'],
  ['M10', 'month-10'],
  ['M11', 'month-11'],
  ['M12', 'month-12'],
]);

const OOXML_DYNAMIC_FILTERS = new Map(
  [...DYNAMIC_FILTERS].map(([ooxml, model]) => [model, ooxml]),
);

export function readXlsxTableFilters(
  table: Document,
  width: number,
): WorkSpreadsheetTableFilter[] {
  const autoFilter = directChild(table.documentElement, 'autoFilter');
  if (!autoFilter) return [];
  return readXlsxAutoFilterColumns(autoFilter, width);
}

export function readXlsxAutoFilterColumns(
  autoFilter: Element,
  width: number,
): WorkSpreadsheetFilter[] {
  const result: WorkSpreadsheetFilter[] = [];
  const observed = new Set<number>();
  for (const column of directChildren(autoFilter, 'filterColumn')) {
    const index = integerAttribute(column, 'colId');
    if (index === null || index < 0 || index >= width || observed.has(index)) {
      continue;
    }
    const criteria = parseFilterCriteria(column);
    if (!criteria) continue;
    observed.add(index);
    result.push({ column: index, criteria });
  }
  return result;
}

export function xlsxTableAutoFilterXml(
  filters: readonly WorkSpreadsheetTableFilter[],
  rangeReference: string,
  totalsRow: boolean,
): string {
  const filterRange = totalsRow
    ? tableRangeWithoutFinalRow(rangeReference)
    : rangeReference;
  return xlsxAutoFilterXml(filters, filterRange);
}

export function xlsxAutoFilterXml(
  filters: readonly WorkSpreadsheetFilter[],
  rangeReference: string,
): string {
  const columns = filters
    .slice()
    .sort((left, right) => left.column - right.column)
    .map(
      (filter) =>
        `<filterColumn colId="${filter.column}">${criteriaXml(
          filter.criteria,
        )}</filterColumn>`,
    )
    .join('');
  return `<autoFilter ref="${escapeXml(rangeReference)}">${columns}</autoFilter>`;
}

function parseFilterCriteria(
  column: Element,
): WorkSpreadsheetFilterCriteria | null {
  const filters = directChild(column, 'filters');
  if (filters) {
    const values = directChildren(filters, 'filter').flatMap((filter) => {
      const value = attribute(filter, 'val');
      return value === null ? [] : [value];
    });
    const includeBlanks = booleanAttribute(filters, 'blank');
    if (!values.length && includeBlanks) return { type: 'blanks' };
    if (values.length) return { type: 'values', values, includeBlanks };
  }
  const customFilters = directChild(column, 'customFilters');
  if (customFilters) {
    const items = directChildren(customFilters, 'customFilter').flatMap(
      (filter) => {
        const value = attribute(filter, 'val');
        return value === null
          ? []
          : [{ operator: attribute(filter, 'operator') ?? 'equal', value }];
      },
    );
    const item = items.length === 1 ? items[0] : undefined;
    if (item) return singleCustomFilter(item);
    if (items.length === 2) {
      const [first, second] = items;
      const conjunction = booleanAttribute(customFilters, 'and');
      if (
        conjunction &&
        first?.operator === 'greaterThanOrEqual' &&
        second?.operator === 'lessThanOrEqual'
      ) {
        return { type: 'between', lower: first.value, upper: second.value };
      }
      if (
        !conjunction &&
        first?.operator === 'lessThan' &&
        second?.operator === 'greaterThan'
      ) {
        return { type: 'not-between', lower: first.value, upper: second.value };
      }
    }
  }
  const top = directChild(column, 'top10');
  if (top) {
    const value = integerAttribute(top, 'val');
    if (value === null || value < 1) return null;
    const upper = attribute(top, 'top') !== '0';
    const percent = booleanAttribute(top, 'percent');
    if ((percent && value > 100) || (!percent && value > 500)) return null;
    if (upper && percent) return { type: 'top-percent', percent: value };
    if (upper) return { type: 'top', count: value };
    if (percent) return { type: 'bottom-percent', percent: value };
    return { type: 'bottom', count: value };
  }
  const dynamic = directChild(column, 'dynamicFilter');
  const kind = dynamic
    ? DYNAMIC_FILTERS.get(attribute(dynamic, 'type') ?? '')
    : null;
  return kind ? { type: 'dynamic', kind } : null;
}

function singleCustomFilter(item: {
  operator: string;
  value: string;
}): WorkSpreadsheetFilterCriteria | null {
  if (item.operator === 'notEqual' && item.value === '') {
    return { type: 'non-blanks' };
  }
  if (item.operator === 'equal') {
    const wildcard = wildcardCriterion(item.value);
    return (
      wildcard ?? {
        type: 'equals',
        value: unescapeWildcards(item.value),
      }
    );
  }
  if (item.operator === 'notEqual') {
    const wildcard = wildcardCriterion(item.value);
    if (!wildcard) {
      return {
        type: 'not-equals',
        value: unescapeWildcards(item.value),
      };
    }
    return wildcard.type === 'contains'
      ? { type: 'does-not-contain', value: wildcard.value }
      : null;
  }
  const types: Record<
    string,
    | 'greater-than'
    | 'greater-than-or-equal'
    | 'less-than'
    | 'less-than-or-equal'
  > = {
    greaterThan: 'greater-than',
    greaterThanOrEqual: 'greater-than-or-equal',
    lessThan: 'less-than',
    lessThanOrEqual: 'less-than-or-equal',
  };
  const type = types[item.operator];
  return type ? { type, value: item.value } : null;
}

function wildcardCriterion(
  source: string,
): { type: 'begins-with' | 'contains' | 'ends-with'; value: string } | null {
  const starts = source.startsWith('*');
  const ends = unescapedFinalAsterisk(source);
  if (!starts && !ends) return null;
  const value = unescapeWildcards(
    source.slice(starts ? 1 : 0, ends ? -1 : undefined),
  );
  if (starts && ends) return { type: 'contains', value };
  return starts ? { type: 'ends-with', value } : { type: 'begins-with', value };
}

function criteriaXml(criteria: WorkSpreadsheetFilterCriteria): string {
  if (criteria.type === 'values') {
    const blank = criteria.includeBlanks ? ' blank="1"' : '';
    return `<filters${blank}>${criteria.values
      .map((value) => `<filter val="${escapeXml(value)}"/>`)
      .join('')}</filters>`;
  }
  if (criteria.type === 'blanks') return '<filters blank="1"/>';
  if (criteria.type === 'non-blanks') {
    return customFilterXml('notEqual', '');
  }
  if (criteria.type === 'between') {
    return customFilterPairXml(
      true,
      'greaterThanOrEqual',
      criteria.lower,
      'lessThanOrEqual',
      criteria.upper,
    );
  }
  if (criteria.type === 'not-between') {
    return customFilterPairXml(
      false,
      'lessThan',
      criteria.lower,
      'greaterThan',
      criteria.upper,
    );
  }
  if (criteria.type === 'top') return topXml(true, false, criteria.count);
  if (criteria.type === 'top-percent')
    return topXml(true, true, criteria.percent);
  if (criteria.type === 'bottom') return topXml(false, false, criteria.count);
  if (criteria.type === 'bottom-percent') {
    return topXml(false, true, criteria.percent);
  }
  if (criteria.type === 'dynamic') {
    const kind = OOXML_DYNAMIC_FILTERS.get(criteria.kind);
    return kind ? `<dynamicFilter type="${kind}"/>` : '';
  }
  const operator = comparisonOperator(criteria.type);
  const value =
    criteria.type === 'contains'
      ? `*${escapeWildcards(criteria.value)}*`
      : criteria.type === 'does-not-contain'
        ? `*${escapeWildcards(criteria.value)}*`
        : criteria.type === 'begins-with'
          ? `${escapeWildcards(criteria.value)}*`
          : criteria.type === 'ends-with'
            ? `*${escapeWildcards(criteria.value)}`
            : criteria.type === 'equals' || criteria.type === 'not-equals'
              ? escapeWildcards(criteria.value)
              : criteria.value;
  return customFilterXml(operator, value);
}

function comparisonOperator(
  type: Exclude<
    WorkSpreadsheetFilterCriteria['type'],
    | 'between'
    | 'blanks'
    | 'bottom'
    | 'bottom-percent'
    | 'dynamic'
    | 'non-blanks'
    | 'not-between'
    | 'top'
    | 'top-percent'
    | 'values'
  >,
): string {
  const operators: Record<typeof type, string> = {
    'begins-with': 'equal',
    contains: 'equal',
    'does-not-contain': 'notEqual',
    'ends-with': 'equal',
    equals: 'equal',
    'greater-than': 'greaterThan',
    'greater-than-or-equal': 'greaterThanOrEqual',
    'less-than': 'lessThan',
    'less-than-or-equal': 'lessThanOrEqual',
    'not-equals': 'notEqual',
  };
  return operators[type];
}

function customFilterXml(operator: string, value: string): string {
  return `<customFilters><customFilter operator="${operator}" val="${escapeXml(
    value,
  )}"/></customFilters>`;
}

function customFilterPairXml(
  conjunction: boolean,
  firstOperator: string,
  firstValue: string,
  secondOperator: string,
  secondValue: string,
): string {
  return `<customFilters and="${conjunction ? 1 : 0}"><customFilter operator="${firstOperator}" val="${escapeXml(
    firstValue,
  )}"/><customFilter operator="${secondOperator}" val="${escapeXml(
    secondValue,
  )}"/></customFilters>`;
}

function topXml(top: boolean, percent: boolean, value: number): string {
  return `<top10 percent="${percent ? 1 : 0}" top="${top ? 1 : 0}" val="${value}"/>`;
}

function tableRangeWithoutFinalRow(reference: string): string {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(reference);
  if (!match) return reference;
  const finalRow = Number(match[4]);
  return `${match[1]}${match[2]}:${match[3]}${Math.max(Number(match[2]), finalRow - 1)}`;
}

function booleanAttribute(element: Element, name: string): boolean {
  return ['1', 'true', 'on'].includes(
    (attribute(element, name) ?? '').toLocaleLowerCase(),
  );
}

function integerAttribute(element: Element, name: string): number | null {
  const source = attribute(element, name);
  if (source === null) return null;
  const value = Number(source);
  return Number.isSafeInteger(value) ? value : null;
}

function escapeWildcards(value: string): string {
  return value
    .replaceAll('~', '~~')
    .replaceAll('*', '~*')
    .replaceAll('?', '~?');
}

function unescapeWildcards(value: string): string {
  return value.replace(/~([~*?])/g, '$1');
}

function unescapedFinalAsterisk(value: string): boolean {
  if (!value.endsWith('*')) return false;
  let escapes = 0;
  for (
    let index = value.length - 2;
    index >= 0 && value[index] === '~';
    index -= 1
  ) {
    escapes += 1;
  }
  return escapes % 2 === 0;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
