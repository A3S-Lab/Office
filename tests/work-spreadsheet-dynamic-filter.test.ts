import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  workSpreadsheetCellIsDate,
  workSpreadsheetDynamicFilterMatcher,
} from '../src/internal/features/work/work-spreadsheet-dynamic-filter';
import type { WorkSpreadsheetDynamicFilter } from '../src/internal/features/work/work-types';

const NOW = new Date(2026, 5, 17, 12, 30);

describe('spreadsheet dynamic AutoFilter evaluation', () => {
  test('evaluates every relative date range with inclusive starts and exclusive ends', () => {
    const scenarios = [
      ['yesterday', ['2026-06-15', '2026-06-16', '2026-06-17'], ['2026-06-16']],
      ['today', ['2026-06-16', '2026-06-17', '2026-06-18'], ['2026-06-17']],
      ['tomorrow', ['2026-06-17', '2026-06-18', '2026-06-19'], ['2026-06-18']],
      [
        'last-week',
        ['2026-06-06', '2026-06-07', '2026-06-13', '2026-06-14'],
        ['2026-06-07', '2026-06-13'],
      ],
      [
        'this-week',
        ['2026-06-13', '2026-06-14', '2026-06-20', '2026-06-21'],
        ['2026-06-14', '2026-06-20'],
      ],
      [
        'next-week',
        ['2026-06-20', '2026-06-21', '2026-06-27', '2026-06-28'],
        ['2026-06-21', '2026-06-27'],
      ],
      [
        'last-month',
        ['2026-04-30', '2026-05-01', '2026-05-31', '2026-06-01'],
        ['2026-05-01', '2026-05-31'],
      ],
      [
        'this-month',
        ['2026-05-31', '2026-06-01', '2026-06-30', '2026-07-01'],
        ['2026-06-01', '2026-06-30'],
      ],
      [
        'next-month',
        ['2026-06-30', '2026-07-01', '2026-07-31', '2026-08-01'],
        ['2026-07-01', '2026-07-31'],
      ],
      [
        'last-quarter',
        ['2025-12-31', '2026-01-01', '2026-03-31', '2026-04-01'],
        ['2026-01-01', '2026-03-31'],
      ],
      [
        'this-quarter',
        ['2026-03-31', '2026-04-01', '2026-06-30', '2026-07-01'],
        ['2026-04-01', '2026-06-30'],
      ],
      [
        'next-quarter',
        ['2026-06-30', '2026-07-01', '2026-09-30', '2026-10-01'],
        ['2026-07-01', '2026-09-30'],
      ],
      [
        'last-year',
        ['2024-12-31', '2025-01-01', '2025-12-31', '2026-01-01'],
        ['2025-01-01', '2025-12-31'],
      ],
      [
        'this-year',
        ['2025-12-31', '2026-01-01', '2026-12-31', '2027-01-01'],
        ['2026-01-01', '2026-12-31'],
      ],
      [
        'next-year',
        ['2026-12-31', '2027-01-01', '2027-12-31', '2028-01-01'],
        ['2027-01-01', '2027-12-31'],
      ],
      [
        'year-to-date',
        ['2025-12-31', '2026-01-01', '2026-06-17', '2026-06-18'],
        ['2026-01-01', '2026-06-17'],
      ],
    ] as const satisfies readonly (readonly [
      WorkSpreadsheetDynamicFilter,
      readonly string[],
      readonly string[],
    ])[];

    for (const [kind, candidates, expected] of scenarios) {
      expect(visibleDates(kind, candidates), kind).toEqual(expected);
    }
  });

  test('matches calendar months and quarters independently of the year', () => {
    for (let month = 1; month <= 12; month += 1) {
      const selected = isoDate(2025, month, 15);
      const selectedInAnotherYear = isoDate(2024, month, 15);
      const other = isoDate(2025, month === 12 ? 1 : month + 1, 15);
      expect(
        visibleDates(`month-${month}` as WorkSpreadsheetDynamicFilter, [
          selected,
          selectedInAnotherYear,
          other,
        ]),
      ).toEqual([selected, selectedInAnotherYear]);
    }

    for (let quarter = 1; quarter <= 4; quarter += 1) {
      const selected = isoDate(2025, (quarter - 1) * 3 + 2, 15);
      const selectedInAnotherYear = isoDate(2024, (quarter - 1) * 3 + 2, 15);
      const other = isoDate(2025, (quarter % 4) * 3 + 2, 15);
      expect(
        visibleDates(`quarter-${quarter}` as WorkSpreadsheetDynamicFilter, [
          selected,
          selectedInAnotherYear,
          other,
        ]),
      ).toEqual([selected, selectedInAnotherYear]);
    }
  });

  test('uses strict arithmetic averages and excludes dates and non-numeric values', () => {
    const cells: Array<Cell | null> = [
      { v: 100 },
      { v: 80 },
      { v: 60 },
      dateCell('2026-06-17'),
      { v: '90' },
      { v: true },
      null,
    ];
    const above = workSpreadsheetDynamicFilterMatcher('above-average', cells, {
      now: NOW,
    });
    const below = workSpreadsheetDynamicFilterMatcher('below-average', cells, {
      now: NOW,
    });

    expect(above).not.toBeNull();
    expect(below).not.toBeNull();
    expect(cells.map((cell) => above?.(cell))).toEqual([
      true,
      false,
      false,
      false,
      false,
      false,
      false,
    ]);
    expect(cells.map((cell) => below?.(cell))).toEqual([
      false,
      false,
      true,
      false,
      false,
      false,
      false,
    ]);

    const extremes: Cell[] = [
      { v: Number.MAX_VALUE },
      { v: Number.MAX_VALUE },
      { v: -Number.MAX_VALUE },
    ];
    const extremeAbove = workSpreadsheetDynamicFilterMatcher(
      'above-average',
      extremes,
      { now: NOW },
    );
    expect(extremes.map((cell) => extremeAbove?.(cell))).toEqual([
      true,
      true,
      false,
    ]);
  });

  test('accepts imported Date values but not unformatted serial numbers', () => {
    const importedDate = {
      v: new Date(2026, 5, 17),
      ct: { fa: 'yyyy-MM-dd', t: 'd' },
    } as unknown as Cell;
    const unformattedSerial: Cell = { v: excelSerial('2026-06-17') };
    const fictionalLeapDay: Cell = {
      v: 60,
      ct: { fa: 'yyyy-MM-dd', t: 'd' },
    };
    const cells = [importedDate, unformattedSerial, fictionalLeapDay];
    const matcher = workSpreadsheetDynamicFilterMatcher('today', cells, {
      now: NOW,
    });

    expect(workSpreadsheetCellIsDate(importedDate)).toBe(true);
    expect(workSpreadsheetCellIsDate(unformattedSerial)).toBe(false);
    expect(cells.map((cell) => matcher?.(cell))).toEqual([true, false, false]);
  });

  test('evaluates typed numeric dates against the workbook 1904 epoch', () => {
    const cells = [
      dateSerialCell(0),
      dateSerialCell(31),
      dateSerialCell(excel1904Serial('2026-06-17')),
    ];
    const january = workSpreadsheetDynamicFilterMatcher('month-1', cells, {
      dateSystem: '1904',
      now: NOW,
    });
    const today = workSpreadsheetDynamicFilterMatcher('today', cells, {
      dateSystem: '1904',
      now: NOW,
    });

    expect(workSpreadsheetCellIsDate(cells[0] ?? null, '1904')).toBe(true);
    expect(workSpreadsheetCellIsDate(cells[0] ?? null)).toBe(false);
    expect(cells.map((cell) => january?.(cell))).toEqual([true, false, false]);
    expect(cells.map((cell) => today?.(cell))).toEqual([false, false, true]);
  });

  test('fails relative date evaluation closed for an invalid clock', () => {
    expect(
      workSpreadsheetDynamicFilterMatcher('today', [dateCell('2026-06-17')], {
        now: new Date(Number.NaN),
      }),
    ).toBeNull();
  });

  test('compiles date filters without materializing or pre-reading the column', () => {
    const unreadCells: Iterable<Cell | null> = {
      [Symbol.iterator]() {
        throw new Error('Date filter compilation eagerly read the column.');
      },
    };
    const matcher = workSpreadsheetDynamicFilterMatcher('today', unreadCells, {
      now: NOW,
    });

    expect(matcher?.(dateCell('2026-06-17'))).toBe(true);
  });
});

function visibleDates(
  kind: WorkSpreadsheetDynamicFilter,
  candidates: readonly string[],
): string[] {
  const cells = candidates.map(dateCell);
  const matcher = workSpreadsheetDynamicFilterMatcher(kind, cells, {
    now: NOW,
  });
  if (!matcher) throw new Error(`Expected a matcher for ${kind}.`);
  return candidates.filter((_, index) => matcher(cells[index] ?? null));
}

function dateCell(value: string): Cell {
  return dateSerialCell(excelSerial(value));
}

function dateSerialCell(value: number): Cell {
  return { v: value, ct: { fa: 'yyyy-MM-dd', t: 'd' } };
}

function excelSerial(value: string): number {
  const [year, month, day] = value.split('-').map(Number);
  return (
    (Date.UTC(year ?? 0, (month ?? 1) - 1, day ?? 1) - Date.UTC(1899, 11, 30)) /
    86_400_000
  );
}

function excel1904Serial(value: string): number {
  const [year, month, day] = value.split('-').map(Number);
  return (
    (Date.UTC(year ?? 0, (month ?? 1) - 1, day ?? 1) - Date.UTC(1904, 0, 1)) /
    86_400_000
  );
}

function isoDate(year: number, month: number, day: number): string {
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}
