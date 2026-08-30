import type { OfficeArtifact, SpreadsheetContent } from '@a3s-lab/office/core';

export const SPREADSHEET_ROW_SORT_FIXTURE = 'spreadsheet-row-sort';
export const SPREADSHEET_ROW_SORT_ARTIFACT_ID = 'fixture-spreadsheet-row-sort';

export function createSpreadsheetRowSortArtifact(): OfficeArtifact {
  const now = Date.now();
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [
    [
      { v: 2, m: '2', bg: '#d9ead3' },
      { v: 1, m: '1', bg: '#d9ead3' },
      { v: 1, m: '1', bg: '#d9ead3' },
      { v: 2, m: '2', bg: '#d9ead3' },
    ],
    [
      { v: 'Gamma', m: 'Gamma', bg: '#fff2cc' },
      { v: 'Alpha', m: 'Alpha' },
      { v: 'Beta', m: 'Beta', fc: '#d84b4f' },
      { v: 'Delta', m: 'Delta' },
    ],
    [
      { v: 'Gamma-G', m: 'Gamma-G', f: '=A2&"-G"' },
      { v: 'Alpha-A', m: 'Alpha-A', f: '=B2&"-A"' },
      { v: 'Beta-B', m: 'Beta-B', f: '=C2&"-B"' },
      { v: 'Delta-D', m: 'Delta-D', f: '=D2&"-D"' },
    ],
  ];

  return {
    id: SPREADSHEET_ROW_SORT_ARTIFACT_ID,
    kind: 'spreadsheet',
    title: 'Left-to-right sort workbook',
    favorite: false,
    createdAt: now,
    updatedAt: now,
    lastOpenedAt: now,
    revision: 1,
    content: {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'row-sort',
          name: 'Horizontal plan',
          status: 1,
          row: 18,
          column: 8,
          data,
          config: {
            columnlen: { 0: 120, 1: 120, 2: 120, 3: 120 },
          },
          luckysheet_select_save: [
            {
              row: [0, 2],
              column: [0, 3],
              row_focus: 0,
              column_focus: 0,
            },
          ],
        },
      ],
    },
  };
}
