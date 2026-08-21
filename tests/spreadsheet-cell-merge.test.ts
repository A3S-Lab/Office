import { describe, expect, test } from '@rstest/core';
import {
  canApplySpreadsheetCellMerge,
  spreadsheetCellMergeApiCalls,
} from '../src/internal/features/work/editors/spreadsheet-cell-merge';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import { createWorkArtifact } from '../src/internal/features/work/work-templates';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

describe('spreadsheet cell merge model', () => {
  test('rejects malformed, native-merge-intersecting, and table ranges', () => {
    const activeSheet = sheet();

    expect(
      canApplySpreadsheetCellMerge(
        activeSheet,
        { row: [0, Number.NaN], column: [0, 1] },
        'merge-cells',
      ),
    ).toBe(false);
    expect(
      canApplySpreadsheetCellMerge(
        activeSheet,
        { row: [1, 2], column: [1, 3] },
        'merge-cells',
      ),
    ).toBe(false);
    expect(
      canApplySpreadsheetCellMerge(
        activeSheet,
        { row: [1, 2], column: [1, 3] },
        'unmerge-cells',
      ),
    ).toBe(true);
    expect(
      canApplySpreadsheetCellMerge(
        {
          ...activeSheet,
          tables: [
            {
              id: 'table-1',
              name: 'Table1',
              range: { row: [4, 6], column: [0, 2] },
              columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }],
              filters: [],
              headerRow: true,
              totalsRow: false,
              style: { family: 'medium', number: 2 },
              showFirstColumn: false,
              showLastColumn: false,
              showRowStripes: true,
              showColumnStripes: false,
            },
          ],
        },
        { row: [5, 5], column: [0, 1] },
        'merge-cells',
      ),
    ).toBe(false);
  });

  test('fills every unmerged cell without mutating the native anchor', () => {
    const activeSheet = sheet();
    const anchor = activeSheet.data?.[1]?.[1];

    const calls = spreadsheetCellMergeApiCalls(
      activeSheet,
      activeSheet.id,
      { row: [1, 2], column: [1, 2] },
      'unmerge-and-fill',
    );

    expect(calls).toEqual([
      {
        name: 'cancelMerge',
        args: [[{ row: [1, 2], column: [1, 2] }], { id: 'sheet-1' }],
      },
      {
        name: 'setCellValuesByRange',
        args: [
          [
            [
              { v: 'North', m: 'North', ht: '0' },
              { v: 'North', m: 'North', ht: '0' },
            ],
            [
              { v: 'North', m: 'North', ht: '0' },
              { v: 'North', m: 'North', ht: '0' },
            ],
          ],
          { row: [1, 2], column: [1, 2] },
          null,
          { id: 'sheet-1' },
        ],
      },
    ]);
    expect(anchor?.mc).toEqual({ r: 1, c: 1, rs: 2, cs: 2 });
  });

  test('preserves native merge ranges through XLSX export and import', async () => {
    const artifact = createWorkArtifact('quarterly-plan');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected the quarterly plan spreadsheet.');
    }
    const firstSheet = artifact.content.sheets[0];
    if (!firstSheet) throw new Error('Expected an active spreadsheet sheet.');
    const content = {
      ...artifact.content,
      sheets: [
        {
          ...firstSheet,
          config: {
            ...firstSheet.config,
            merge: {
              ...firstSheet.config?.merge,
              '0_0': { r: 0, c: 0, rs: 2, cs: 3 },
            },
          },
        },
        ...artifact.content.sheets.slice(1),
      ],
    };

    const blob = await createWorkArtifactBlob({ ...artifact, content });
    const imported = await importWorkFile(
      new File([blob], 'quarterly-plan-merged.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected an imported spreadsheet.');
    }

    expect(imported.content.sheets[0]?.config?.merge?.['0_0']).toEqual({
      r: 0,
      c: 0,
      rs: 2,
      cs: 3,
    });
  });
});

function sheet(): WorkSpreadsheetSheet {
  return {
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 20,
    column: 10,
    config: {
      merge: {
        '1_1': { r: 1, c: 1, rs: 2, cs: 2 },
      },
    },
    data: [
      [],
      [
        null,
        {
          v: 'North',
          m: 'North',
          ht: '0',
          mc: { r: 1, c: 1, rs: 2, cs: 2 },
        },
        { mc: { r: 1, c: 1 } },
      ],
      [null, { mc: { r: 1, c: 1 } }, { mc: { r: 1, c: 1 } }],
    ],
  };
}
