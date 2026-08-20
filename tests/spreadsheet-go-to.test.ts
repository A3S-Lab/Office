import { describe, expect, test } from '@rstest/core';
import { resolveSpreadsheetGoToTarget } from '../src/internal/features/work/editors/spreadsheet-go-to';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet Go To resolution', () => {
  test('resolves local and quoted cross-sheet A1 ranges', () => {
    const content = goToContent();

    expect(
      resolveSpreadsheetGoToTarget(content, 'sheet-1', '$D$8:$B$4'),
    ).toEqual({
      ok: true,
      target: {
        displayReference: "'Q1''s Sales'!B4:D8",
        range: { row: [3, 7], column: [1, 3] },
        selection: {
          row: [3, 7],
          column: [1, 3],
          row_focus: 3,
          column_focus: 1,
        },
        sheetId: 'sheet-1',
        source: 'reference',
      },
    });
    expect(
      resolveSpreadsheetGoToTarget(
        content,
        'sheet-1',
        "'Archive 2025'!$C$9:$E$12",
      ),
    ).toMatchObject({
      ok: true,
      target: {
        displayReference: "'Archive 2025'!C9:E12",
        range: { row: [8, 11], column: [2, 4] },
        sheetId: 'sheet-2',
        source: 'reference',
      },
    });
  });

  test('prefers the active worksheet name and resolves workbook names', () => {
    const content = goToContent();

    expect(
      resolveSpreadsheetGoToTarget(content, 'sheet-1', 'Pipeline'),
    ).toMatchObject({
      ok: true,
      target: {
        displayReference: "'Q1''s Sales'!B4:D8",
        name: 'Pipeline',
        range: { row: [3, 7], column: [1, 3] },
        sheetId: 'sheet-1',
        source: 'named-range',
      },
    });
    expect(
      resolveSpreadsheetGoToTarget(content, 'sheet-1', 'ArchiveBlock'),
    ).toMatchObject({
      ok: true,
      target: {
        displayReference: "'Archive 2025'!C9:E12",
        name: 'ArchiveBlock',
        sheetId: 'sheet-2',
        source: 'named-range',
      },
    });
    expect(
      resolveSpreadsheetGoToTarget(
        content,
        'sheet-1',
        "'Archive 2025'!Pipeline",
      ),
    ).toMatchObject({
      ok: true,
      target: {
        displayReference: "'Archive 2025'!A20:B21",
        name: 'Pipeline',
        sheetId: 'sheet-2',
        source: 'named-range',
      },
    });
  });

  test('rejects unsafe, hidden, ambiguous, unsupported, and out-of-bounds targets', () => {
    const content = goToContent();
    const cases = [
      ['', 'empty'],
      ['A1,B2', 'multiple-ranges'],
      ["'Missing'!A1", 'sheet-not-found'],
      ["'Hidden'!A1", 'hidden-sheet'],
      ['M1', 'out-of-bounds'],
      ['A41', 'out-of-bounds'],
      ['XFE1', 'out-of-bounds'],
      ['FormulaName', 'unsupported-name'],
      ['UnknownName', 'invalid-reference'],
      ['[External.xlsx]Sheet1!A1', 'sheet-not-found'],
    ] as const;

    for (const [value, code] of cases) {
      expect(
        resolveSpreadsheetGoToTarget(content, 'sheet-1', value),
      ).toMatchObject({ ok: false, code });
    }
  });

  test('does not materialize maximum-dimension sparse worksheets', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sparse',
          name: 'Sparse',
          status: 1,
          row: 1_048_576,
          column: 16_384,
          celldata: [{ r: 1_048_575, c: 16_383, v: { v: 'edge' } }],
        },
      ],
    } satisfies WorkSpreadsheetContent;

    expect(
      resolveSpreadsheetGoToTarget(content, 'sparse', 'XFD1048576'),
    ).toMatchObject({
      ok: true,
      target: {
        range: {
          row: [1_048_575, 1_048_575],
          column: [16_383, 16_383],
        },
      },
    });
    expect(content.sheets[0]).not.toHaveProperty('data');
  });
});

function goToContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: "Q1's Sales",
        status: 1,
        row: 40,
        column: 12,
      },
      {
        id: 'sheet-2',
        name: 'Archive 2025',
        row: 100,
        column: 20,
      },
      {
        id: 'hidden',
        name: 'Hidden',
        hide: 1,
        row: 40,
        column: 12,
      },
    ],
    namedRanges: [
      {
        id: 'active-pipeline',
        name: 'Pipeline',
        reference: '$B$4:$D$8',
        scopeSheetId: 'sheet-1',
      },
      {
        id: 'archive-pipeline',
        name: 'Pipeline',
        reference: '$A$20:$B$21',
        scopeSheetId: 'sheet-2',
      },
      {
        id: 'workbook-pipeline',
        name: 'Pipeline',
        reference: "'Archive 2025'!$F$2:$G$4",
      },
      {
        id: 'archive-block',
        name: 'ArchiveBlock',
        reference: "'Archive 2025'!$C$9:$E$12",
      },
      {
        id: 'formula-name',
        name: 'FormulaName',
        reference: '=SUM(A1:A2)',
      },
    ],
  };
}
